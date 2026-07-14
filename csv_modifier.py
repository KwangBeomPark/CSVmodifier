import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import pandas as pd
import os
import re
import sys
from datetime import datetime
import decimal
from collections import Counter

__version__ = "1.5.0"

class ParsedNumber:
    __slots__ = ['value', 'orig_decimals', 'orig_text']

    def __init__(self, value: decimal.Decimal, orig_decimals: int, orig_text: str):
        self.value = value
        self.orig_decimals = orig_decimals
        self.orig_text = orig_text


# Pre-compiled patterns (compiling once matters a lot on large files).
# _DATE_HINT_RE cheaply rejects non-date cells so we skip the expensive
# strptime attempts on the vast majority of values.
_DATE_HINT_RE = re.compile(r'\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}')
# Every character Excel/Windows can smuggle into a cell as a line break:
# CR/LF plus vertical tab (Alt+Enter survives copy-paste as \x0b), form feed,
# NEL, and the Unicode line/paragraph separators.
_LINEBREAK_RE = re.compile('[\r\n\x0b\x0c\x85\u2028\u2029]+')
_EN_NUM_RE = re.compile(r'-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?')
_PL_NUM_RE = re.compile(r'-?(?:\d+|\d{1,3}(?: \d{3})+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?')

_NUMBER_FORMAT_OPTIONS = (
    "English · 1,234.56",
    "Polish · 1 234,56",
)
_OUTPUT_FORMAT_OPTIONS = (
    "CSV 텍스트 파일 (.csv)",
    "Excel 통합 문서 (.xlsx)",
)

class CSVModifierApp:
    @staticmethod
    def _resource_path(filename):
        """Find bundled assets both during development and in PyInstaller builds."""
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_dir, filename)

    def __init__(self, root):
        self.root = root
        self.root.title(f"CSV Modifier v{__version__}")
        self.root.geometry("820x760")
        self.root.minsize(700, 650)

        try:
            self.root.iconbitmap(self._resource_path("icon.ico"))
        except Exception:
            pass

        self._app_icon_image = None
        self._header_icon_image = None
        try:
            self._app_icon_image = tk.PhotoImage(file=self._resource_path("icon.png"))
            self._header_icon_image = self._app_icon_image.subsample(20, 20)
            self.root.iconphoto(True, self._app_icon_image)
        except tk.TclError:
            pass

        # A calm, high-contrast palette that remains readable with the clam theme.
        page_bg = "#F4F7FB"
        surface = "#FFFFFF"
        navy = "#102A43"
        text = "#243B53"
        muted = "#627D98"
        border = "#D9E2EC"
        accent = "#147D8A"
        accent_active = "#0F6974"

        self.root.configure(background=page_bg)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=("Segoe UI", 10), background=page_bg, foreground=text)
        style.configure("App.TFrame", background=page_bg)
        style.configure("Header.TFrame", background=navy)
        style.configure("Header.Icon.TLabel", background=navy)
        style.configure("Header.Title.TLabel", background=navy, foreground="#FFFFFF", font=("Segoe UI Semibold", 20))
        style.configure("Header.Subtitle.TLabel", background=navy, foreground="#C9D6E2", font=("Segoe UI", 10))
        style.configure("Card.TLabelframe", background=surface, bordercolor=border, relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", background=surface, foreground=navy, font=("Segoe UI Semibold", 10))
        style.configure("TLabel", background=surface, foreground=text)
        style.configure("Muted.TLabel", background=surface, foreground=muted, font=("Segoe UI", 9))
        style.configure("Help.TLabel", background=surface, foreground=muted, font=("Segoe UI", 8))
        style.configure("Footer.TLabel", background=page_bg, foreground=muted, font=("Segoe UI", 9))
        style.configure("Field.TLabel", background=surface, foreground=text, font=("Segoe UI Semibold", 9))
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=text, padding=(8, 6), bordercolor=border)
        style.map("TEntry", bordercolor=[("focus", accent)], fieldbackground=[("readonly", "#F8FAFC")])
        style.configure("TCombobox", fieldbackground="#FFFFFF", foreground=text, padding=(6, 5), bordercolor=border)
        style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")], bordercolor=[("focus", accent)])
        style.configure("Primary.TButton", background=accent, foreground="#FFFFFF", font=("Segoe UI Semibold", 10), padding=(18, 10), borderwidth=0)
        style.map("Primary.TButton", background=[("active", accent_active), ("pressed", accent_active), ("disabled", "#9FB3C8")], foreground=[("disabled", "#F0F4F8")])
        style.configure("Secondary.TButton", background="#E8F1F5", foreground=accent, font=("Segoe UI Semibold", 9), padding=(12, 8), borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#D4E8EE"), ("pressed", "#C4DDE5")])
        style.configure("Status.TFrame", background=navy)
        style.configure("Status.TLabel", background=navy, foreground="#D9E2EC", font=("Segoe UI", 9))
        style.configure("App.Horizontal.TProgressbar", troughcolor=border, background=accent, bordercolor=border, lightcolor=accent, darkcolor=accent)

        main = ttk.Frame(root, style="App.TFrame", padding=(24, 20, 24, 18))
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(5, weight=1)

        header = ttk.Frame(main, style="Header.TFrame", padding=(22, 18))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        title_column = 0
        if self._header_icon_image is not None:
            ttk.Label(header, image=self._header_icon_image, style="Header.Icon.TLabel").grid(
                row=0, column=0, rowspan=2, sticky="w", padx=(0, 14)
            )
            title_column = 1
        ttk.Label(header, text="CSV Modifier", style="Header.Title.TLabel").grid(row=0, column=title_column, sticky="w")
        ttk.Label(
            header,
            text="파일을 정리하고 숫자 표기를 맞춘 새 파일을 만듭니다.",
            style="Header.Subtitle.TLabel",
        ).grid(row=1, column=title_column, sticky="w", pady=(3, 0))

        file_section = ttk.LabelFrame(main, text="1. 원본 파일 선택", style="Card.TLabelframe", padding=(18, 14))
        file_section.grid(row=1, column=0, sticky="ew", pady=(16, 12))
        file_section.columnconfigure(0, weight=1)

        # File Path
        self.filepath = tk.StringVar()
        self.lbl_file = ttk.Entry(file_section, textvariable=self.filepath, state="readonly")
        self.lbl_file.grid(row=0, column=0, sticky="ew")
        ttk.Button(file_section, text="파일 찾기...", command=self.browse_file, style="Secondary.TButton").grid(row=0, column=1, padx=(10, 0))
        ttk.Label(
            file_section,
            text="CSV, TXT, Excel(.xlsx/.xlsm)을 고르세요. 원본 파일은 바꾸지 않습니다.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        settings = ttk.Frame(main, style="App.TFrame")
        settings.grid(row=2, column=0, sticky="nsew")
        settings.columnconfigure(0, weight=1)
        settings.columnconfigure(1, weight=1)

        import_section = ttk.LabelFrame(settings, text="2. 파일 읽는 방법", style="Card.TLabelframe", padding=(18, 14))
        import_section.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        import_section.columnconfigure(1, weight=1)

        output_section = ttk.LabelFrame(settings, text="3. 저장 방법", style="Card.TLabelframe", padding=(18, 14))
        output_section.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        output_section.columnconfigure(1, weight=1)

        # Delimiter
        self.delimiter = tk.StringVar(value=",")
        ttk.Label(import_section, text="파일 구분 기호", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_delimiter = ttk.Entry(import_section, textvariable=self.delimiter, width=8)
        self.ent_delimiter.grid(row=0, column=1, sticky="ew")
        self._delimiter_user_set = False
        self.ent_delimiter.bind("<KeyRelease>", self._on_delimiter_user_change)
        ttk.Label(
            import_section,
            text="쉼표(,), 세미콜론(;), 탭(\\t) 중 파일 안의 표를 나누는 기호입니다.",
            style="Help.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 12))

        # Number Format
        self.num_format = tk.StringVar(value=_NUMBER_FORMAT_OPTIONS[0])
        ttk.Label(import_section, text="숫자 표기 방식", style="Field.TLabel").grid(row=2, column=0, sticky="w")
        self.combo_format = ttk.Combobox(
            import_section,
            textvariable=self.num_format,
            values=_NUMBER_FORMAT_OPTIONS,
            state="readonly",
            width=22,
        )
        self.combo_format.grid(row=2, column=1, sticky="ew")
        self.number_format_help = tk.StringVar()
        ttk.Label(import_section, textvariable=self.number_format_help, style="Help.TLabel", wraplength=310).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        self.combo_format.bind("<<ComboboxSelected>>", self._update_number_format_help)

        # Max Columns
        self.max_cols = tk.StringVar(value="0")
        ttk.Label(output_section, text="표의 열 개수", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_max_cols = ttk.Entry(output_section, textvariable=self.max_cols, width=8)
        self.ent_max_cols.grid(row=0, column=1, sticky="ew")
        ttk.Label(
            output_section,
            text="파일을 고르면 자동으로 채워집니다. 표의 실제 열 수와 다를 때만 바꾸세요.",
            style="Help.TLabel",
            wraplength=310,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 12))

        # Output Format
        self.out_format = tk.StringVar(value=_OUTPUT_FORMAT_OPTIONS[0])
        ttk.Label(output_section, text="저장 파일 형식", style="Field.TLabel").grid(row=2, column=0, sticky="w")
        self.combo_out_format = ttk.Combobox(
            output_section,
            textvariable=self.out_format,
            values=_OUTPUT_FORMAT_OPTIONS,
            state="readonly",
            width=22,
        )
        self.combo_out_format.grid(row=2, column=1, sticky="ew")
        self.output_format_help = tk.StringVar()
        ttk.Label(output_section, textvariable=self.output_format_help, style="Help.TLabel", wraplength=310).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        self.combo_out_format.bind("<<ComboboxSelected>>", self._update_output_hint)

        action_area = ttk.Frame(main, style="App.TFrame")
        action_area.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        action_area.columnconfigure(0, weight=1)
        self.output_hint = tk.StringVar(value="저장 파일: 파일을 고르면 예상 이름을 보여 드립니다.")
        ttk.Label(action_area, textvariable=self.output_hint, style="Footer.TLabel").grid(row=0, column=0, sticky="w")

        # Process Button
        self.btn_process = ttk.Button(action_area, text="정리하고 저장하기", command=self.process_csv, style="Primary.TButton")
        self.btn_process.grid(row=0, column=1, sticky="e")

        # Progress bar (advances during processing)
        self.progress = ttk.Progressbar(main, mode="determinate", maximum=100, style="App.Horizontal.TProgressbar")
        self.progress.grid(row=4, column=0, sticky="ew", pady=(14, 0))

        result_section = ttk.LabelFrame(
            main,
            text="처리 결과 · 무엇이 정리되었나요?",
            style="Card.TLabelframe",
            padding=(12, 10),
        )
        result_section.grid(row=5, column=0, sticky="nsew", pady=(14, 0))
        result_section.columnconfigure(0, weight=1)
        result_section.rowconfigure(0, weight=1)
        self.result_text = tk.Text(
            result_section,
            height=8,
            wrap="word",
            relief="flat",
            borderwidth=0,
            background="#FFFFFF",
            foreground=text,
            font=("Segoe UI", 9),
            padx=6,
            pady=4,
            state="disabled",
        )
        self.result_text.grid(row=0, column=0, sticky="nsew")

        # Status bar
        self.log_text = tk.StringVar(value="준비됨 · 파일을 고르고 숫자 표기 방식을 선택하세요.")
        status_bar = ttk.Frame(root, style="Status.TFrame", padding=(24, 9))
        status_bar.grid(row=1, column=0, sticky="ew")
        ttk.Label(status_bar, textvariable=self.log_text, style="Status.TLabel").grid(row=0, column=0, sticky="w")

        self._update_number_format_help()
        self._update_output_hint()
        self._set_result_text("아직 처리한 파일이 없습니다. 파일을 고르고 ‘정리하고 저장하기’를 누르세요.")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="CSV, TXT 또는 Excel 파일 선택",
            filetypes=(
                ("CSV/TXT/Excel files", "*.csv *.txt *.xlsx *.xlsm"),
                ("All files", "*.*"),
            )
        )
        if filename:
            self.filepath.set(filename)
            detected_delim = self._detect_delimiter(filename)
            self._apply_detected_delimiter(detected_delim)
            self.update_max_columns()
            self._update_output_hint()

    @staticmethod
    def _number_mode(selection):
        """Map friendly combobox text to the parser's stable internal mode."""
        return 'Polish' if str(selection).startswith('Polish') else 'English'

    @staticmethod
    def _output_format(selection):
        """Map friendly combobox text to the stable output format identifier."""
        return 'Excel (.xlsx)' if str(selection).startswith('Excel') else 'CSV'

    def _update_number_format_help(self, event=None):
        if self._number_mode(self.num_format.get()) == 'Polish':
            message = "Polish: 1 234,56 · 공백/점은 천 단위, 쉼표는 소수점입니다."
        else:
            message = "English: 1,234.56 · 쉼표는 천 단위, 점은 소수점입니다."
        self.number_format_help.set(message)

    @staticmethod
    def _output_extension(output_fmt):
        return '.xlsx' if output_fmt == 'Excel (.xlsx)' else '.csv'

    @classmethod
    def _build_output_path(cls, file_path, output_fmt, timestamp=None):
        """Create a timestamped output path and avoid overwriting a same-minute run."""
        stamp = (timestamp or datetime.now()).strftime('%Y%m%d_%H%M')
        extension = cls._output_extension(output_fmt)
        directory = os.path.dirname(file_path)
        base_name = f"processed_output_{stamp}"
        candidate = os.path.join(directory, f"{base_name}{extension}")
        sequence = 2
        while os.path.exists(candidate):
            candidate = os.path.join(directory, f"{base_name}_{sequence:02d}{extension}")
            sequence += 1
        return candidate

    def _update_output_hint(self, event=None):
        output_fmt = self._output_format(self.out_format.get())
        extension = self._output_extension(output_fmt)
        expected_name = f"processed_output_YYYYMMDD_HHMM{extension}"
        self.output_format_help.set(
            "CSV는 텍스트 파일입니다. Excel은 바로 열어 계산할 수 있습니다."
            if output_fmt == 'CSV'
            else "Excel 파일로 저장합니다. 큰 숫자는 정확도를 위해 텍스트로 보존될 수 있습니다."
        )
        if self.filepath.get():
            self.output_hint.set(f"저장 위치: 원본과 같은 폴더 · 이름: {expected_name}")
        else:
            self.output_hint.set(f"저장 파일 이름: {expected_name}")

    def _set_result_text(self, message):
        """Show the processing explanation inside the app instead of a success popup."""
        result_widget = getattr(self, 'result_text', None)
        if result_widget is None:
            return
        result_widget.configure(state='normal')
        result_widget.delete('1.0', tk.END)
        result_widget.insert('1.0', message)
        result_widget.configure(state='disabled')

    @staticmethod
    def _is_excel(file_path):
        return os.path.splitext(file_path)[1].lower() in ('.xlsx', '.xlsm', '.xls')

    def _on_delimiter_user_change(self, event=None):
        """Remember an explicit delimiter choice so file selection cannot overwrite it."""
        self._delimiter_user_set = True
        self.update_max_columns(event)

    def _apply_detected_delimiter(self, detected_delim):
        """Apply an inferred delimiter only while the user has not chosen one."""
        if not detected_delim or getattr(self, '_delimiter_user_set', False):
            return False
        self.delimiter.set('\\t' if detected_delim == '\t' else detected_delim)
        return True

    def _detect_delimiter(self, file_path, sample_rows=50):
        """Infer a delimiter from logical CSV records, not physical text lines.

        This deliberately tolerates introductory one-column rows and quoted
        multi-line cells. The candidate with the most stable non-trivial row
        width wins; a one-column candidate is never selected automatically.
        """
        if self._is_excel(file_path):
            return None
        try:
            enc = self._detect_encoding(file_path)
            candidates = [';', ',', '\t', '|']
            best_delimiter = None
            best_score = None
            for cand in candidates:
                with open(file_path, 'r', encoding=enc, newline='') as f:
                    reader = csv.reader(f, delimiter=cand)
                    rows = []
                    for row in reader:
                        if row:
                            rows.append(row)
                        if len(rows) >= sample_rows:
                            break
                if not rows:
                    continue

                widths = [len(row) for row in rows]
                width, count = max(Counter(widths).items(), key=lambda item: (item[1], item[0]))
                if width <= 1:
                    continue

                # Ratio first prevents a one-off delimiter in prose from
                # beating a consistently structured file; width breaks ties.
                score = (count / len(widths), count, width)
                if best_score is None or score > best_score:
                    best_score = score
                    best_delimiter = cand
            return best_delimiter
        except Exception:
            return None

    @staticmethod
    def _normalize_delim(delim):
        # Let the user type "\t" or "tab" for tab-separated files
        # (Excel's "Unicode Text" export is tab-separated).
        return {'\\t': '\t', 'tab': '\t', 'TAB': '\t', 'Tab': '\t'}.get(delim, delim)

    def _detect_encoding(self, file_path):
        """Pick the most plausible text encoding.

        UTF-8 (with or without BOM) is authoritative when it decodes cleanly.
        Among legacy codecs, prefer the one that yields the fewest suspicious
        characters (C1 controls / replacement chars), which are a strong signal
        of a wrong codepage. This avoids cp1252 (or cp949) silently mis-decoding
        a file just because it happens to accept the bytes.
        """
        with open(file_path, 'rb') as f:
            raw = f.read()

        # Excel's "Unicode Text" / "CSV UTF-16" exports are UTF-16. The BOM is
        # authoritative, and it must be checked before UTF-8: ASCII-only
        # UTF-16 also decodes "successfully" as UTF-8 (with NULs between
        # every character), which then crashes csv with "line contains NUL".
        if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
            return 'utf-16'
        if raw and raw.count(b'\x00') > len(raw) // 4:
            for enc in ('utf-16-le', 'utf-16-be'):
                try:
                    raw.decode(enc)
                    return enc
                except UnicodeDecodeError:
                    pass

        for enc in ('utf-8-sig', 'utf-8'):
            try:
                raw.decode(enc)
                return enc
            except UnicodeDecodeError:
                pass

        best = None  # (encoding, suspicious_char_count)
        for enc in ('cp949', 'cp1252'):
            try:
                text = raw.decode(enc)
            except UnicodeDecodeError:
                continue
            score = sum(1 for ch in text if ('\x80' <= ch <= '\x9f') or ch == '�')
            if best is None or score < best[1]:
                best = (enc, score)

        if best is not None:
            return best[0]
        return 'latin-1'  # never fails; last-resort so the app degrades gracefully

    def read_file_lines(self, file_path, delim, max_lines=None):
        if self._is_excel(file_path):
            return self._read_excel_rows(file_path, max_lines), 'Excel'
        delim = self._normalize_delim(delim)
        if not delim or len(delim) != 1:
            raise ValueError("Delimiter must be a single character.")
        enc = self._detect_encoding(file_path)
        with open(file_path, 'r', encoding=enc, newline='') as f:
            reader = csv.reader(f, delimiter=delim)
            rows = []
            for i, row in enumerate(reader):
                if max_lines is not None and i >= max_lines:
                    break
                rows.append(row)
        return rows, enc

    def _read_excel_rows(self, file_path, max_lines=None):
        """Read an Excel sheet into rows of native values (no CSV round-trip,
        so quoting/encoding problems cannot occur). Trailing empty cells are
        trimmed per row so garbage/header detection works the same as for CSV."""
        df = pd.read_excel(file_path, header=None, dtype=object, nrows=max_lines)
        rows = []
        for rec in df.itertuples(index=False, name=None):
            row = []
            for v in rec:
                if v is None or (not isinstance(v, str) and pd.isna(v)):
                    row.append('')
                elif isinstance(v, datetime):
                    row.append(v.date() if (v.hour, v.minute, v.second) == (0, 0, 0) else v)
                else:
                    row.append(v)
            while row and row[-1] == '':
                row.pop()
            rows.append(row)
        return rows

    @staticmethod
    def _repair_split_rows(rows, max_c):
        """Rejoin records that unquoted in-cell line breaks split across
        physical lines (a frequent defect in Excel-exported or hand-edited
        CSVs). A short row is merged with the following row(s) only when the
        pieces reassemble to exactly max_c columns; genuinely short rows are
        left untouched and padded later as before."""
        out = []
        repaired = 0
        i = 0
        n = len(rows)
        while i < n:
            row = rows[i]
            if not 0 < len(row) < max_c:
                out.append(row)
                i += 1
                continue
            # Case 1: a break inside the LAST field leaves the record itself
            # complete and spills the remainder onto its own one-cell line —
            # glue that remainder onto the previous row's last cell. This must
            # be checked first: forward-merging such an orphan would swallow
            # the next real record instead.
            if len(row) == 1 and out and len(out[-1]) == max_c:
                out[-1][-1] = f"{out[-1][-1]} {row[0]}".strip()
                repaired += 1
                i += 1
                continue
            # Case 2: a break inside a MIDDLE field leaves every fragment
            # short. Rebuild the record from consecutive short fragments,
            # joining the two halves of the broken cell with a space. Never
            # consume a complete row, and commit only if the pieces
            # reassemble to exactly max_c columns.
            merged = list(row)
            j = i + 1
            while j < n and len(merged) < max_c and j - i <= 20:
                nxt = rows[j]
                if len(nxt) >= max_c or len(merged) + max(len(nxt), 1) - 1 > max_c:
                    break
                if nxt:
                    merged[-1] = f"{merged[-1]} {nxt[0]}".strip()
                    merged.extend(nxt[1:])
                j += 1
            if len(merged) == max_c and j > i + 1:
                out.append(merged)
                repaired += 1
                i = j
                continue
            out.append(row)
            i += 1
        return out, repaired

    def update_max_columns(self, event=None):
        file_path = self.filepath.get()
        delim = self.delimiter.get()
        if not file_path or not os.path.exists(file_path):
            return
        if not delim:
            return
        
        try:
            rows, enc = self.read_file_lines(file_path, delim, max_lines=10)
            max_cols = max((len(r) for r in rows), default=0)
            
            self.max_cols.set(str(max_cols))
            self.log_text.set(f"Detected Max Columns: {max_cols} (Encoding: {enc})")
        except Exception as e:
            self.log_text.set("Error detecting columns (Encoding/Delimiter issue)")

    @staticmethod
    def parse_number(val, mode):
        """Parse a safely validated English or Polish number into Decimal.

        A parsed value retains its source decimal scale so CSV and Excel
        exporters can preserve values such as ``500,00`` cell by cell.
        Invalid or identifier-like values deliberately remain text.
        """
        if not isinstance(val, str):
            return val
        v = val.strip()
        if not v:
            return val

        v_norm = v.replace('\u00A0', ' ').replace('\u202F', ' ').replace('\u2212', '-')

        sign = ''
        if v_norm.startswith('-') or v_norm.startswith('+'):
            sign = v_norm[0]
            v_norm = v_norm[1:]

        if not v_norm:
            return val

        if mode == 'Polish':
            if not re.fullmatch(r'[0-9 \.,]+', v_norm):
                return val

            parts = v_norm.split(',')
            if len(parts) > 2:
                return val

            int_part = parts[0]
            frac_part = parts[1] if len(parts) == 2 else None

            if frac_part == '':
                return val

            if not int_part and frac_part:
                int_part = '0'

            if ' ' in int_part and '.' in int_part:
                return val

            sep = ' ' if ' ' in int_part else '.' if '.' in int_part else None
            if sep:
                groups = int_part.split(sep)
                if not groups[0] or len(groups[0]) > 3:
                    return val
                for g in groups[1:]:
                    if len(g) != 3:
                        return val
                int_part = ''.join(groups)

            if frac_part and ('.' in frac_part or ' ' in frac_part):
                return val

            if len(int_part) > 1 and int_part.startswith('0'):
                return val

            orig_decimals = len(frac_part) if frac_part else 0
            dec_str = f"{sign}{int_part}.{frac_part}" if frac_part else f"{sign}{int_part}"

            try:
                d = decimal.Decimal(dec_str)
                return ParsedNumber(d, orig_decimals, val)
            except decimal.InvalidOperation:
                return val

        elif mode == 'English':
            if not re.fullmatch(r'[0-9\.,]+', v_norm):
                return val

            parts = v_norm.split('.')
            if len(parts) > 2:
                return val

            int_part = parts[0]
            frac_part = parts[1] if len(parts) == 2 else None

            if frac_part == '':
                return val

            if not int_part and frac_part:
                int_part = '0'

            if ',' in int_part:
                groups = int_part.split(',')
                if not groups[0] or len(groups[0]) > 3:
                    return val
                for g in groups[1:]:
                    if len(g) != 3:
                        return val
                int_part = ''.join(groups)

            if frac_part and ',' in frac_part:
                return val

            if len(int_part) > 1 and int_part.startswith('0'):
                return val

            orig_decimals = len(frac_part) if frac_part else 0
            dec_str = f"{sign}{int_part}.{frac_part}" if frac_part else f"{sign}{int_part}"

            try:
                d = decimal.Decimal(dec_str)
                return ParsedNumber(d, orig_decimals, val)
            except decimal.InvalidOperation:
                return val

        return val

    @staticmethod
    def _decimal_significant_digits(value):
        """Count significant decimal digits without treating trailing scale zeros as precision."""
        normalized = value.normalize()
        digits = normalized.as_tuple().digits
        return 1 if all(digit == 0 for digit in digits) else len(digits)

    @classmethod
    def _requires_excel_text(cls, value):
        """Excel stores at most 15 significant decimal digits reliably."""
        return cls._decimal_significant_digits(value.value) > 15

    @staticmethod
    def _format_csv_value(value, decimal_separator):
        """Render a converted value without losing its original decimal scale."""
        if isinstance(value, bool):
            return value
        if isinstance(value, ParsedNumber):
            text = f"{value.value:.{value.orig_decimals}f}"
            return text.replace('.', decimal_separator) if decimal_separator != '.' else text
        if isinstance(value, float):
            if pd.isna(value):
                return ''
            text = repr(value)
            return text.replace('.', decimal_separator) if decimal_separator != '.' else text
        return value

    @classmethod
    def _write_excel_output(cls, df, out_path):
        """Write converted data and return how many values were protected as text."""
        import openpyxl

        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(list(df.columns))
        large_numbers_as_text = 0

        for row in df.itertuples(index=False):
            output_row = []
            for value in row:
                if isinstance(value, ParsedNumber):
                    if cls._requires_excel_text(value):
                        large_numbers_as_text += 1
                        output_row.append(value.orig_text)
                    else:
                        output_row.append(float(value.value))
                else:
                    output_row.append(value)
            worksheet.append(output_row)

        for row_index, row in enumerate(df.itertuples(index=False), start=2):
            for column_index, value in enumerate(row, start=1):
                if isinstance(value, ParsedNumber) and not cls._requires_excel_text(value):
                    cell = worksheet.cell(row=row_index, column=column_index)
                    cell.number_format = '0' if value.orig_decimals == 0 else '0.' + ('0' * value.orig_decimals)

        workbook.save(out_path)
        return large_numbers_as_text

    def parse_date(self, val):
        """Try to parse a string into a date. Returns a date on success, else the original value."""
        if not isinstance(val, str):
            return val
        v = val.strip()
        # Skip strptime unless the value even looks like a date. This is the
        # single biggest speed-up on large files (most cells are not dates).
        if not v or not _DATE_HINT_RE.fullmatch(v):
            return val
        formats = [
            "%Y-%m-%d", "%Y/%m/%d",
            "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y",
            "%m/%d/%Y",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(v, fmt).date()
            except ValueError:
                continue
        return val

    def convert_value(self, val, mode):
        """Type coercion for a single cell: try date, then number, else keep original text."""
        d = self.parse_date(val)
        if not isinstance(d, str):
            return d
        return self.parse_number(val, mode)

    def _set_progress(self, pct, msg=None):
        """Move the progress bar and (optionally) the status text, then repaint."""
        try:
            self.progress['value'] = max(0, min(100, pct))
            if msg is not None:
                self.log_text.set(msg)
            self.root.update_idletasks()
        except Exception:
            pass

    def process_csv(self):
        file_path = self.filepath.get()
        delim = self._normalize_delim(self.delimiter.get())
        mode = self._number_mode(self.num_format.get())
        output_fmt = self._output_format(self.out_format.get())

        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("파일을 선택해 주세요", "정리할 CSV, TXT 또는 Excel 파일을 선택해 주세요.")
            return

        is_excel = self._is_excel(file_path)
        if not is_excel and (not delim or len(delim) != 1):
            messagebox.showerror("구분 기호 확인", "구분 기호는 한 글자여야 합니다. 탭은 \\t로 입력하세요.")
            return

        try:
            max_c = int(self.max_cols.get())
        except ValueError:
            messagebox.showerror("열 개수 확인", "표의 열 개수에는 숫자를 입력해 주세요.")
            return

        if max_c <= 0:
            messagebox.showerror("열 개수 확인", "표의 열 개수는 1 이상이어야 합니다.")
            return

        self._set_progress(0, "파일을 읽는 중...")
        self.root.update()

        try:
            rows, enc = self.read_file_lines(file_path, delim)
            self._set_progress(10, f"{len(rows):,}개 행을 읽었습니다. 표 구조를 확인하는 중...")

            # Find the start index (first row that has max_cols) to skip garbage
            start_idx = 0
            for i, row in enumerate(rows):
                if len(row) == max_c:
                    start_idx = i
                    break

            garbage_skipped = start_idx
            data_rows = rows[start_idx:]

            # Rejoin records that unquoted in-cell line breaks split across
            # physical lines (Excel input needs no repair: cells arrive intact).
            rows_repaired = 0
            if not is_excel:
                data_rows, rows_repaired = self._repair_split_rows(data_rows, max_c)

            # Pad or truncate short/long rows to exactly max_c columns
            padded_rows = []
            for r in data_rows:
                r = list(r)
                if len(r) < max_c:
                    r.extend([''] * (max_c - len(r)))
                elif len(r) > max_c:
                    r = r[:max_c]
                padded_rows.append(r)

            if not padded_rows:
                messagebox.showwarning("처리할 표를 찾지 못했습니다", "선택한 열 개수와 맞는 행이 없습니다.")
                self.log_text.set("처리할 표를 찾지 못했습니다.")
                return

            # Promote the first matching row to the header, remaining rows are data
            header_row = padded_rows[0]
            body_rows = padded_rows[1:]

            if not body_rows:
                messagebox.showwarning("데이터가 없습니다", "열 이름은 찾았지만 정리할 데이터 행이 없습니다.")
                self.log_text.set("열 이름 다음에 데이터가 없습니다.")
                return

            # Build unique, non-empty column names from the header
            col_names = []
            seen = {}
            for i, name in enumerate(header_row):
                name = '' if name is None else str(name)
                name = _LINEBREAK_RE.sub(' ', name).strip()
                name = name or f"Column{i+1}"
                if name in seen:
                    seen[name] += 1
                    name = f"{name}_{seen[name]}"
                else:
                    seen[name] = 0
                col_names.append(name)

            # Create DataFrame from the data rows (header excluded)
            df = pd.DataFrame(body_rows, columns=col_names)

            # Type coercion (date -> number -> text) in a single pass per column
            # that also tallies what changed and records each column's original
            # decimal precision (used to keep values like "500,00" intact on export).
            dec_sep = ',' if mode == 'Polish' else '.'
            stats = {'numbers': 0, 'dates': 0, 'flattened': 0, 'large_numbers_as_text': 0}
            ncols = len(df.columns)

            for ci, col in enumerate(df.columns):
                def convert(x):
                    if not isinstance(x, str):
                        return x  # Excel input arrives already typed (dates, numbers)
                    # Flatten in-cell line breaks — including the exotic ones Excel
                    # produces (\x0b from Alt+Enter copy-paste, U+2028/U+2029) — so a
                    # multi-line field becomes a single physical line in the output.
                    if _LINEBREAK_RE.search(x):
                        x = _LINEBREAK_RE.sub(' ', x)
                        stats['flattened'] += 1
                    d = self.parse_date(x)
                    if not isinstance(d, str):
                        stats['dates'] += 1
                        return d
                    n = self.parse_number(x, mode)
                    if isinstance(n, ParsedNumber):
                        stats['numbers'] += 1
                        return n
                    return n

                df[col] = df[col].map(convert)
                self._set_progress(10 + int(70 * (ci + 1) / ncols),
                                   f"숫자와 날짜를 정리하는 중... ({ci + 1}/{ncols}개 열)")

            # Replace NaNs with empty strings
            df = df.fillna('')

            # Export
            self._set_progress(85, "새 파일을 저장하는 중...")
            out_path = self._build_output_path(file_path, output_fmt)
            if output_fmt == "Excel (.xlsx)":
                stats['large_numbers_as_text'] = self._write_excel_output(df, out_path)
            else:
                sep = ';' if mode == 'Polish' else ','
                out_df = df.copy()
                for col in out_df.columns:
                    out_df[col] = out_df[col].map(
                        lambda value: self._format_csv_value(value, dec_sep)
                    )

                out_df.to_csv(out_path, sep=sep, index=False, encoding='utf-8-sig')

            self._set_progress(100, "완료되었습니다.")

            summary = (
                f"저장 완료 · {os.path.basename(out_path)}\n"
                f"저장 위치: {os.path.dirname(out_path)}\n\n"
                "이번에 정리한 내용\n"
                f"• 데이터 행 {len(body_rows):,}개와 열 {ncols}개를 새 파일로 저장했습니다.\n"
                f"• 위쪽의 표가 아닌 행 {garbage_skipped:,}개를 건너뛰고, 첫 번째 정상 행을 열 이름으로 사용했습니다.\n"
                f"• 숫자 표기 {stats['numbers']:,}개와 날짜 {stats['dates']:,}개를 읽기 쉬운 값으로 정리했습니다.\n"
                f"• 셀 안의 줄바꿈 {stats['flattened']:,}개를 한 줄 텍스트로 바꿨습니다.\n"
                f"• 여러 줄로 끊어진 행 {rows_repaired:,}개를 다시 이었습니다.\n"
                f"• Excel 정확도 보호를 위해 큰 숫자 {stats['large_numbers_as_text']:,}개를 텍스트로 보존했습니다.\n"
                f"• 읽은 파일 인코딩: {enc}"
            )
            self.log_text.set(
                f"완료 · {len(body_rows):,}개 행 저장 → {os.path.basename(out_path)}"
            )
            self._set_result_text(summary)

        except Exception as e:
            messagebox.showerror("처리 중 오류", f"파일을 정리하지 못했습니다.\n\n{str(e)}")
            self.log_text.set("처리 중 오류가 발생했습니다.")
        finally:
            self._set_progress(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVModifierApp(root)
    root.mainloop()
