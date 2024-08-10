# CSV to Excel Converter

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE?logo=powershell&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

A **PowerShell batch converter** that transforms all CSV files in an input directory into formatted Excel (.xlsx) files. Processed CSV files are automatically renamed with a timestamp to prevent reprocessing.

## Features

- **Batch Processing** — Converts all `.csv` files in the input directory
- **Auto-Sizing** — Excel columns are automatically sized for readability
- **Logging** — Timestamped logs for each conversion run
- **File Archiving** — Processed CSV files are renamed with a timestamp suffix
- **Semicolon Delimiter** — Supports European-style CSV files (`;` separator)

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Language | PowerShell 5.1+ |
| Excel Module | ImportExcel |

## Installation

```bash
git clone https://github.com/jmanu1983/csv-to-excel-converter.git
cd csv-to-excel-converter
```

Install the required PowerShell module:

```powershell
Install-Module -Name ImportExcel -Force
```

## Usage

1. Place your CSV files in the `input/` directory
2. Run the script:

```powershell
.\convertCSVtoExcel.ps1
```

3. Find the generated Excel files in `output/`

## Project Structure

```
csv-to-excel-converter/
├── convertCSVtoExcel.ps1   # Main conversion script
├── input/                  # Place CSV files here
├── output/                 # Generated Excel files
├── logs/                   # Conversion logs
└── README.md
```

## License

This project is licensed under the MIT License.
