import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import pandas as pd
import os
import re
from datetime import datetime

__version__ = "1.1.0"

class CSVModifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"CSV Modifier v{__version__}")
        self.root.geometry("720x560")
        self.root.minsize(600, 500)

        try:
            self.root.iconbitmap(os.path.join(os.path.dirname(__file__), "icon.ico"))
        except Exception:
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
        style.configure("Header.Title.TLabel", background=navy, foreground="#FFFFFF", font=("Segoe UI Semibold", 20))
        style.configure("Header.Subtitle.TLabel", background=navy, foreground="#C9D6E2", font=("Segoe UI", 10))
        style.configure("Card.TLabelframe", background=surface, bordercolor=border, relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", background=surface, foreground=navy, font=("Segoe UI Semibold", 10))
        style.configure("TLabel", background=surface, foreground=text)
        style.configure("Muted.TLabel", background=surface, foreground=muted, font=("Segoe UI", 9))
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

        main = ttk.Frame(root, style="App.TFrame", padding=(24, 20, 24, 18))
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        header = ttk.Frame(main, style="Header.TFrame", padding=(22, 18))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="CSV Modifier", style="Header.Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Clean, normalize, and export tabular data with confidence.",
            style="Header.Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))

        file_section = ttk.LabelFrame(main, text="1. Choose a source file", style="Card.TLabelframe", padding=(18, 14))
        file_section.grid(row=1, column=0, sticky="ew", pady=(16, 12))
        file_section.columnconfigure(0, weight=1)

        # File Path
        self.filepath = tk.StringVar()
        self.lbl_file = ttk.Entry(file_section, textvariable=self.filepath, state="readonly")
        self.lbl_file.grid(row=0, column=0, sticky="ew")
        ttk.Button(file_section, text="Browse...", command=self.browse_file, style="Secondary.TButton").grid(row=0, column=1, padx=(10, 0))
        ttk.Label(file_section, text="CSV and TXT files are supported.", style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        settings = ttk.Frame(main, style="App.TFrame")
        settings.grid(row=2, column=0, sticky="nsew")
        settings.columnconfigure(0, weight=1)
        settings.columnconfigure(1, weight=1)

        import_section = ttk.LabelFrame(settings, text="2. Configure import", style="Card.TLabelframe", padding=(18, 14))
        import_section.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        import_section.columnconfigure(1, weight=1)

        output_section = ttk.LabelFrame(settings, text="3. Choose output", style="Card.TLabelframe", padding=(18, 14))
        output_section.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        output_section.columnconfigure(1, weight=1)

        # Delimiter
        self.delimiter = tk.StringVar(value=",")
        ttk.Label(import_section, text="Delimiter", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_delimiter = ttk.Entry(import_section, textvariable=self.delimiter, width=8)
        self.ent_delimiter.grid(row=0, column=1, sticky="ew", pady=(0, 12))
        self.ent_delimiter.bind("<KeyRelease>", self.update_max_columns)

        # Number Format
        self.num_format = tk.StringVar(value="English")
        ttk.Label(import_section, text="Number format", style="Field.TLabel").grid(row=1, column=0, sticky="w")
        self.combo_format = ttk.Combobox(import_section, textvariable=self.num_format, values=["English", "Polish"], state="readonly", width=15)
        self.combo_format.grid(row=1, column=1, sticky="ew")

        # Max Columns
        self.max_cols = tk.StringVar(value="0")
        ttk.Label(output_section, text="Max columns", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_max_cols = ttk.Entry(output_section, textvariable=self.max_cols, width=8)
        self.ent_max_cols.grid(row=0, column=1, sticky="ew", pady=(0, 12))

        # Output Format
        self.out_format = tk.StringVar(value="CSV")
        ttk.Label(output_section, text="Output format", style="Field.TLabel").grid(row=1, column=0, sticky="w")
        self.combo_out_format = ttk.Combobox(output_section, textvariable=self.out_format, values=["CSV", "Excel (.xlsx)"], state="readonly", width=15)
        self.combo_out_format.grid(row=1, column=1, sticky="ew")

        action_area = ttk.Frame(main, style="App.TFrame")
        action_area.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        action_area.columnconfigure(0, weight=1)
        ttk.Label(action_area, text="Your original file is never changed.", style="Footer.TLabel").grid(row=0, column=0, sticky="w")

        # Process Button
        self.btn_process = ttk.Button(action_area, text="Process & Save", command=self.process_csv, style="Primary.TButton")
        self.btn_process.grid(row=0, column=1, sticky="e")

        # Status bar
        self.log_text = tk.StringVar(value="Ready.")
        status_bar = ttk.Frame(root, style="Status.TFrame", padding=(24, 9))
        status_bar.grid(row=1, column=0, sticky="ew")
        ttk.Label(status_bar, textvariable=self.log_text, style="Status.TLabel").grid(row=0, column=0, sticky="w")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select CSV/TXT File",
            filetypes=(("CSV/TXT files", "*.csv *.txt"), ("All files", "*.*"))
        )
        if filename:
            self.filepath.set(filename)
            self.update_max_columns()

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

    def parse_number(self, val, mode):
        if not isinstance(val, str):
            return val
        v = val.strip()
        if not v:
            return val
        if mode == 'English':
            if not re.fullmatch(r'-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?', v):
                return val
            num_str = v.replace(',', '')
        elif mode == 'Polish':
            if not re.fullmatch(r'-?(?:\d+|\d{1,3}(?: \d{3})+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?', v):
                return val
            num_str = v.replace(' ', '').replace('.', '').replace(',', '.')
        else:
            return val

        if '.' not in num_str:
            return int(num_str)
        integer_part, fractional_part = num_str.split('.', 1)
        if set(fractional_part) == {'0'}:
            return int(integer_part)
        return float(num_str)

    def parse_date(self, val):
        """Try to parse a string into a date. Returns a date on success, else the original value."""
        if not isinstance(val, str):
            return val
        v = val.strip()
        if not v:
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

    def process_csv(self):
        file_path = self.filepath.get()
        delim = self.delimiter.get()
        mode = self.num_format.get()
        output_fmt = self.out_format.get()
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file.")
            return

        if not delim or len(delim) != 1:
            messagebox.showerror("Error", "Delimiter must be a single character.")
            return

        try:
            max_c = int(self.max_cols.get())
        except ValueError:
            messagebox.showerror("Error", "Max Columns must be a number.")
            return

        if max_c <= 0:
            messagebox.showerror("Error", "Max Columns must be greater than 0.")
            return

        self.log_text.set("Processing...")
        self.root.update()

        try:
            rows, enc = self.read_file_lines(file_path, delim)

            # Find the start index (first row that has max_cols) to skip garbage
            start_idx = 0
            for i, row in enumerate(rows):
                if len(row) == max_c:
                    start_idx = i
                    break
            
            data_rows = rows[start_idx:]

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
                messagebox.showwarning("Warning", "No rows matched the max columns criteria.")
                self.log_text.set("Warning: No valid rows found.")
                return

            # Promote the first matching row to the header, remaining rows are data
            header_row = padded_rows[0]
            body_rows = padded_rows[1:]

            if not body_rows:
                messagebox.showwarning("Warning", "Header found but no data rows to process.")
                self.log_text.set("Warning: No data rows after header.")
                return

            # Build unique, non-empty column names from the header
            col_names = []
            seen = {}
            for i, name in enumerate(header_row):
                name = (name or '').strip() or f"Column{i+1}"
                if name in seen:
                    seen[name] += 1
                    name = f"{name}_{seen[name]}"
                else:
                    seen[name] = 0
                col_names.append(name)

            # Create DataFrame from the data rows (header excluded)
            df = pd.DataFrame(body_rows, columns=col_names)

            # Apply type coercion (date -> number -> text)
            for col in df.columns:
                df[col] = df[col].apply(lambda x: self.convert_value(x, mode))

            # Replace NaNs with empty strings
            df = df.fillna('')

            # Export
            if output_fmt == "Excel (.xlsx)":
                # Excel keeps real numeric/date types so the values stay computable.
                out_path = os.path.join(os.path.dirname(file_path), 'processed_output.xlsx')
                df.to_excel(out_path, index=False)
            else:
                out_path = os.path.join(os.path.dirname(file_path), 'processed_output.csv')
                sep = ';' if mode == 'Polish' else ','
                dec_sep = ',' if mode == 'Polish' else '.'

                # Preserve each column's original decimal precision so values like
                # Polish "500,00" are written back as "500,00" instead of "500,0".
                col_decimals = [0] * len(col_names)
                for r in body_rows:
                    for ci in range(len(col_names)):
                        cell = r[ci].strip()
                        if dec_sep in cell and isinstance(self.parse_number(cell, mode), float):
                            col_decimals[ci] = max(col_decimals[ci],
                                                   len(cell.rsplit(dec_sep, 1)[-1]))

                out_df = df.copy()
                for ci, col in enumerate(out_df.columns):
                    width = col_decimals[ci]

                    def fmt(x, width=width):
                        if isinstance(x, bool):
                            return x
                        if isinstance(x, float):
                            if pd.isna(x):
                                return ''
                            return f"{x:.{width}f}".replace('.', dec_sep) if dec_sep != '.' else f"{x:.{width}f}"
                        return x

                    out_df[col] = out_df[col].apply(fmt)

                out_df.to_csv(out_path, sep=sep, index=False, encoding='utf-8-sig')

            self.log_text.set(f"Saved to: {out_path}")
            messagebox.showinfo("Success", f"File processed successfully!\nSaved to: {out_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            self.log_text.set("Error during processing.")

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVModifierApp(root)
    root.mainloop()
