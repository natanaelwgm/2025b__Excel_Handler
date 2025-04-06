# Excel Handler: An AI Capabilities Experiment

This repository was created as an experiment to compare and evaluate the capabilities of different AI models (Gemini and ChatGPT/Claude) when tasked with similar programming challenges, specifically Excel file manipulation tasks. The two implementations in this repo demonstrate how different AI models approach the same problem.

## About this Experiment

This project contains two parallel implementations of the same functionality:

- **gemini/** directory: Code generated with Google's Gemini AI
- **o1pro/** directory: Code generated with OpenAI's ChatGPT/Claude

The goal was to compare how these different AI systems approach similar tasks, their coding styles, algorithmic choices, and overall effectiveness in generating functional code for Excel manipulation tasks.

## Features

- Extract data from Excel files, including cell values and formulas
- Compare two Excel files to identify all differences (values, formulas, structure)
- Generate detailed comparison reports
- Create sample Excel files for demonstration purposes

## Requirements

- Python 3.6+
- openpyxl library
- pytz library

## Installation

### Option 1: Install directly from GitHub

```bash
pip install git+https://github.com/natanaelwgm/2025b__Excel_Handler.git
```

### Option 2: Clone and install locally

1. Clone the repository:
```bash
git clone https://github.com/natanaelwgm/2025b__Excel_Handler.git
cd 2025b__Excel_Handler
```

2. Create a virtual environment and activate it:
```bash
# On macOS/Linux
python -m venv venv
source venv/bin/activate

# On Windows
python -m venv venv
venv\Scripts\activate
```

3. Install the package in development mode:
```bash
pip install -e .
```

## Usage

### Using as an imported package

```python
# Using the gemini implementation
from excel_handler.gemini import read_excel_file_data, compare_excel_data

# Read Excel files
data1 = read_excel_file_data("file1.xlsx")
data2 = read_excel_file_data("file2.xlsx")

# Compare the data
differences = compare_excel_data(data1, data2)
print(differences)

# Or using the o1pro implementation
from excel_handler.o1pro import read_excel_to_dict, compare_excel_dicts

excel1_data = read_excel_to_dict("file1.xlsx")
excel2_data = read_excel_to_dict("file2.xlsx")
comparison = compare_excel_dicts(excel1_data, excel2_data)
```

### Running the demo

```bash
# Run the gemini implementation demo
python -c "from excel_handler.gemini import main; main()"

# Run the o1pro implementation demo
python -c "from excel_handler.o1pro import main; main()"
```

## Available Modules

This package contains two different implementations of the same functionality:

1. `excel_handler.gemini` - The Gemini AI implementation
   - `read_excel_file_data(filepath)`: Reads all data from an Excel file
   - `compare_excel_data(data1, data2)`: Compares data from two Excel files
   - `create_demo_excel_file(filepath, file_data)`: Creates a demo Excel file with specified data

2. `excel_handler.o1pro` - The ChatGPT/Claude implementation
   - `read_excel_to_dict(filepath)`: Reads all data from an Excel file
   - `compare_excel_dicts(dict1, dict2)`: Compares data from two Excel files
   - `export_excel_dict_to_txt(excel_dict, file_path, source_excel_path)`: Exports Excel data to a text file
   - `export_comparison_to_txt(diff_dict, dict1, dict2, file_path, excel1_name, excel2_name)`: Exports comparison to a text file

## Key Observations

Some interesting differences between the two AI implementations:

- **Function Naming**: Different naming conventions for similar functions
- **Data Structures**: Subtle differences in how data is organized
- **Error Handling**: Variations in error handling approaches
- **Documentation**: Different documentation styles and level of detail
- **Algorithm Design**: Different approaches to solving the same problems

## Example Output

When comparing two Excel files, the tool provides a detailed report of differences:
- Sheets that exist only in one file
- Cells with different values or formulas
- Summary statistics of differences

## License

MIT License 