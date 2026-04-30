# Course Intersection Analyzer

A desktop application for Windows that reads student enrollment data from **Excel** (`.xlsx` / `.xls`) or from a **JSON API**, then finds students who take **both** of two courses you select.

Within students enrolled in the **first** course, it lists those who also take the **second** course.

## Features

- **Excel**: Pick a file and map columns interactively.
- **API**: Enter a URL and fetch JSON records (see [API response format](#api-response-format)).
- **Flexible columns**: Map student ID (optional), **first name**, **last name**, and course name separately—suites exports where name fields are split.
- **Matching**: Prefer student ID; if missing or you choose “No student ID”, match by combined **First name + Last name** (case-insensitive normalization).
- **Results**: Table with student ID, full name, and the two relevant courses; optional export to Excel.
- **Executable**: Build a single-file `.exe` with PyInstaller (see [Building an executable](#building-an-executable)).

## Requirements

- Python 3.10+
- Dependencies (see `requirements.txt`): `pandas`, `openpyxl`, `customtkinter`, `requests`, `pyinstaller` (for builds)

## Installation

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

The window title is **Ders Kesisim Analizi** (Turkish UI labels); functionality is language-independent regarding your data column names.

## Usage

1. **Choose data source**
   - **Excel**: Set source to Excel, click **Excel Dosyasi Sec**, choose `.xlsx` or `.xls`.
   - **API**: Set source to API, enter the URL, click **API Verisini Cek**.
2. **Map columns**
   - **Student number** (optional): pick the column, or **Ogrenci no yok** if you have no ID column.
   - **First name** and **Last name**: pick the corresponding columns.
   - **Course name**: pick the course column.
   - Click **Sutunlari Onayla** to load unique course names.
3. **Select courses**
   - Choose **first** and **second** course from the dropdowns (must be different).
4. **Analyze**
   - Click **Analiz Et**. Results appear in the table and a short summary line above it.
5. **Export** (optional)
   - Click **Sonucu Excel'e Aktar** to save the filtered list.

## Expected data shape

Each row should represent one enrollment line: one student, one course (same student may appear on multiple rows).

Example columns (your headers can differ; you map them in the UI):

- Student ID — optional  
- First name  
- Last name  
- Course name  

## API response format

The app expects valid JSON in one of these forms:

- A **JSON array** of objects (each object is one row), e.g. `[{ "StudentNo": "1001", "First": "Ali", ... }, ...]`
- An object with a **`data`** key holding that array: `{ "data": [ ... ] }`

Empty payloads or non-tabular JSON will show a clear error in the UI.

## Sample data

| File | Purpose |
|------|---------|
| `sample_data/ornek_veri.csv` | Sample rows with separate first/last name columns |
| `sample_data/ornek_api_data.json` | Same idea as JSON for API testing |
| `sample_data/create_sample_excel.py` | Converts the CSV to `sample_data/ornek_veri.xlsx` |

```bash
python sample_data/create_sample_excel.py
```

For a quick API test, serve `ornek_api_data.json` with any static file server and paste that URL into the app.

## Building an executable

On Windows, from the project root:

```bat
build_exe.bat
```

Equivalent PyInstaller command:

```bash
pyinstaller --onefile --windowed --name DersKesisimAnalizi main.py
```

If `assets/app_icon.ico` exists, `build_exe.bat` adds `--icon=assets/app_icon.ico`.

## Project layout

```text
├── main.py
├── requirements.txt
├── README.md
├── LICENSE
├── build_exe.bat
├── app/
│   ├── gui.py
│   ├── excel_reader.py
│   ├── analyzer.py
│   └── exporter.py
├── assets/
│   └── app_icon.ico
└── sample_data/
    ├── ornek_veri.csv
    ├── ornek_api_data.json
    └── create_sample_excel.py
```

## Credits

Prepared by **Tarık Aykan** · [betanova.tech](https://betanova.tech)

## License

This repository includes an MIT `LICENSE` file.
