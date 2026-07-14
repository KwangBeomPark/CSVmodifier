import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import pandas as pd
import os

class CSVModifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Modifier")
        self.root.geometry("500x400")
        
        # File Path
        self.filepath = tk.StringVar()
        tk.Label(root, text="Target File:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.lbl_file = tk.Entry(root, textvariable=self.filepath, width=40, state='readonly')
        self.lbl_file.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(root, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=10, pady=10)
        
        # Delimiter
        self.delimiter = tk.StringVar(value=",")
        tk.Label(root, text="Delimiter:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.ent_delimiter = tk.Entry(root, textvariable=self.delimiter, width=10)
        self.ent_delimiter.grid(row=1, column=1, sticky="w", padx=10)
        self.ent_delimiter.bind("<KeyRelease>", self.update_max_columns)
        
        # Number Format
        self.num_format = tk.StringVar(value="English")
        tk.Label(root, text="Number Format:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.combo_format = ttk.Combobox(root, textvariable=self.num_format, values=["English", "Polish"], state="readonly", width=15)
        self.combo_format.grid(row=2, column=1, sticky="w", padx=10)
        
        # Max Columns
        self.max_cols = tk.StringVar(value="0")
        tk.Label(root, text="Max Columns:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.ent_max_cols = tk.Entry(root, textvariable=self.max_cols, width=10)
        self.ent_max_cols.grid(row=3, column=1, sticky="w", padx=10)

        # Output Format
        self.out_format = tk.StringVar(value="CSV")
        tk.Label(root, text="Output Format:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.combo_out_format = ttk.Combobox(root, textvariable=self.out_format, values=["CSV", "Excel (.xlsx)"], state="readonly", width=15)
        self.combo_out_format.grid(row=4, column=1, sticky="w", padx=10)
        
        # Process Button
        self.btn_process = tk.Button(root, text="Process & Save", command=self.process_csv, bg="green", fg="white", font=("Arial", 10, "bold"))
        self.btn_process.grid(row=5, column=0, columnspan=3, pady=20)
        
        # Log
        self.log_text = tk.StringVar(value="Ready.")
        tk.Label(root, textvariable=self.log_text, fg="blue").grid(row=6, column=0, columnspan=3, pady=5)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select CSV/TXT File",
            filetypes=(("CSV/TXT files", "*.csv *.txt"), ("All files", "*.*"))
        )
        if filename:
            self.filepath.set(filename)
            self.update_max_columns()

    def read_file_lines(self, file_path, delim, max_lines=None):
        encodings = ['utf-8-sig', 'cp949', 'cp1252']
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc, newline='') as f:
                    reader = csv.reader(f, delimiter=delim)
                    rows = []
                    for i, row in enumerate(reader):
                        if max_lines is not None and i >= max_lines:
                            break
                        rows.append(row)
                return rows, enc
            except UnicodeDecodeError:
                continue
        raise ValueError("Could not decode file with standard encodings.")

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
        try:
            if mode == 'English':
                num_str = v.replace(',', '')
                return float(num_str)
            elif mode == 'Polish':
                num_str = v.replace(' ', '').replace('.', '').replace(',', '.')
                return float(num_str)
        except ValueError:
            return val
        return val

    def process_csv(self):
        file_path = self.filepath.get()
        delim = self.delimiter.get()
        mode = self.num_format.get()
        output_fmt = self.out_format.get()
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file.")
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
            
            # Pad or truncate short/long rows
            padded_rows = []
            for r in data_rows:
                if len(r) < max_c:
                    r.extend([''] * (max_c - len(r)))
                elif len(r) > max_c:
                    r = r[:max_c]
                padded_rows.append(r)

            if not padded_rows:
                messagebox.showwarning("Warning", "No rows matched the max columns criteria.")
                self.log_text.set("Warning: No valid rows found.")
                return

            # Create DataFrame
            col_names = [f"Column{i+1}" for i in range(max_c)]
            df = pd.DataFrame(padded_rows, columns=col_names)

            # Apply number parsing
            for col in df.columns:
                df[col] = df[col].apply(lambda x: self.parse_number(x, mode))
                
            # Replace NaNs with empty strings
            df = df.fillna('')

            # Export
            if output_fmt == "Excel (.xlsx)":
                out_path = os.path.join(os.path.dirname(file_path), 'processed_output.xlsx')
                df.to_excel(out_path, index=False)
            else:
                out_path = os.path.join(os.path.dirname(file_path), 'processed_output.csv')
                if mode == 'Polish':
                    df.to_csv(out_path, sep=';', decimal=',', index=False, encoding='utf-8-sig')
                else:
                    df.to_csv(out_path, sep=',', decimal='.', index=False, encoding='utf-8-sig')

            self.log_text.set(f"Saved to: {out_path}")
            messagebox.showinfo("Success", f"File processed successfully!\nSaved to: {out_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            self.log_text.set("Error during processing.")

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVModifierApp(root)
    root.mainloop()
