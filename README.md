# Excel Handler

A Python utility for working with Excel files. This tool can read Excel files (preserving both values and formulas), compare Excel files to identify differences, and generate demo files for testing.

## Features

- Extract data from Excel files, including cell values and formulas
- Compare two Excel files and identify all differences (values, formulas, structure)
- Generate detailed comparison reports
- Create sample Excel files for demonstration purposes

## Requirements

- Python 3.6+
- openpyxl library

## Installation

1. Clone the repository:
```
git clone https://github.com/yourusername/excel-handler.git
cd excel-handler
```

2. Create a virtual environment and activate it:
```
# On macOS/Linux
python -m venv venv
source venv/bin/activate

# On Windows
python -m venv venv
venv\Scripts\activate
```

3. Install the required dependencies:
```
pip install -r requirements.txt
```

## Usage

Run the main script to see a demonstration of the Excel Handler's capabilities:

```
python excel_handler.py
```

### Key Functions

- `read_excel_file_data(filepath)`: Reads all data from an Excel file
- `compare_excel_data(data1, data2)`: Compares data from two Excel files
- `create_demo_excel_file(filepath, file_data)`: Creates a demo Excel file with specified data

## Example Output

When comparing two Excel files, the tool provides a detailed report of differences:
- Sheets that exist only in one file
- Cells with different values or formulas
- Summary statistics of differences

## License

[Specify your license here] 