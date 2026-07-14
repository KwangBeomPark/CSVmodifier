*Read this in other languages: [English](README.md), [한국어](README.ko.md)*

# CSV Modifier

CSV Modifier is an intuitive, easy-to-use CSV / TXT / Excel conversion and cleaning program designed for users without any coding experience.
It is specifically built to solve the following common issues:

## Key Features
- **Safe Data Parsing**: Even if text (strings) contain line breaks (Enter), it reads them as a single cell without errors, just like Excel or Power Query. When exporting, internal cell line breaks (including special characters like vertical tab `\x0b`, U+2028, etc.) are replaced with spaces so that one record is stored in exactly one line.
- **Broken Multi-line Record Recovery**: For files where line breaks occur inside a cell without quotes causing a single record to split across multiple lines (a common defect in Excel exports), the program automatically pieces the fragments back together. The number of recovered rows is displayed in the result summary.
- **Direct Excel File Input**: Reads `.xlsx` / `.xlsm` files directly without converting them to CSV first. Cell values are imported with their original types, inherently preventing encoding or quoting issues.
- **English and Polish Number Format Conversion**: 
  - Cleanly converts Polish number formats (e.g., `1 234,56` or `1.234,56`) to valid floating-point numbers.
  - English number formats (e.g., `1,234.56`) are also supported and can be selected.
- **Automatic Variable Column Detection & Garbage Data Removal**: Analyzes the top rows of the file to infer the maximum number of columns. It ignores top garbage rows (Garbage Header) that lack enough columns, promotes the first valid row to the **actual header**, and extracts only the data.
- **Automatic Date/Number Type Conversion**: Converts date formats (e.g., `2026-07-14`, `14.07.2026`) to dates, and numbers to integers/floats. If conversion fails, it retains the original text. It also preserves the original decimal places (e.g., `500,00`).
- **Output Format Selection**: 
  - Standard CSV format (uses semicolon `;` delimiter when Polish format is selected).
  - Direct export to Excel file (`.xlsx`).
- **Automatic Encoding Detection**: Prioritizes UTF-8 detection and automatically detects UTF-16 (with or without BOM), which is Excel's "Unicode Text" format. For legacy encodings (`cp949`, `cp1252`), it selects the one with the fewest broken characters, ensuring Korean Excel files are processed without text corruption. For tab-separated files, simply enter `\t` as the delimiter.
- **High-Speed Processing for Large Files**: Optimizes date/number detection to process over 100,000 rows in seconds.
- **Progress Display and Result Summary**: Shows a progress bar and steps during processing. Provides a summary of converted number/date cells, column count, and encoding upon completion.

## Download & How to Run
You can download the standalone executable file (`.exe`) from the link below and use it immediately. No need to install Python or any additional libraries.

👉 **[Download Latest Version (v1.4.0)](https://github.com/KwangBeomPark/CSVmodifier/releases/tag/v1.4.0)**

1. Go to the link and download the `csv_modifier.exe` file.
2. Double-click the executable to open it.
3. Select the file to convert, specify the delimiter, number format, and output format, then click the **[Process & Save]** button.
4. The `processed_output.csv` (or `.xlsx`) will be created in the same folder as the original file.