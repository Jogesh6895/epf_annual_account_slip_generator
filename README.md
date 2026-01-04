# EPF Report Generator

A comprehensive tool for generating Employees' Provident Fund (EPF) annual account slips from Excel data.

## Features

- **Excel Input/Output Support**: Read data from Excel files, output to both CSV and Excel formats
- **Comprehensive Validation**: Validates sheet dimensions, column counts, and row counts before processing
- **EPF Calculations**:
  - Employee contribution: 12% of wages
  - EPS contribution: 8.33% of wages
  - Employer contribution: 3.67% of wages
  - Monthly balance tracking
  - Interest calculation based on average balance
  - Withdrawal processing
- **Error Handling**: Custom exceptions with clear error messages
- **Type Hints**: Full type annotations for better code clarity
- **Modular Design**: Separated functions for data loading, validation, calculation, and output
- **Formatted Output**: Excel output with headers, colors, and borders
- **Performance Tracking**: Displays execution time

## Installation

### Prerequisites
- Python 3.6 or higher
- openpyxl library

### Install Dependencies
```bash
pip install openpyxl
```

## Usage

### Running the Calculator

```bash
cd epf_report_generator
python3 epf_calculator.py
```

### Quick Start with Sample Data

```bash
cd epf_report_generator
# Sample files are already in InputFiles/
# Just run the calculator
python3 epf_calculator.py
```

The provided `Input.xlsx` and `Sample_Input.xlsx` contain 5 sample employees with realistic data:
- John Doe, Jane Smith, Raj Kumar, Sunita Sharma, Amit Patel
- Wage range: 12,000 - 25,000
- Various opening balances and withdrawals
- Ready to run and test immediately

### Input File Requirements

#### Sample Input File
A sample input file with 5 dummy employees is provided for testing:
- `InputFiles/Sample_Input.xlsx` - Sample template for reference
- `InputFiles/Input.xlsx` - Copy of sample file, ready to use

To create your own input file, create an Excel file named `InputFiles/Input.xlsx` with the following sheets:

#### Sheet: Wages
- **Columns**: 14 columns
  - Column 1: Account Number (A/C No.)
  - Column 2: Employee Name
  - Columns 3-14: Monthly wages (April to March, 12 months)
- **Row 1**: Header row
- **Rows 2 onwards**: Employee data

#### Sheet: OB_EE
- **Columns**: 1 column
  - Column 1: Opening balance for Employee share
- **Row 1**: Header row
- **Rows 2 onwards**: One row per employee

#### Sheet: OB_ER
- **Columns**: 1 column
  - Column 1: Opening balance for Employer share
- **Row 1**: Header row
- **Rows 2 onwards**: One row per employee

#### Sheet: OB_EPS
- **Columns**: 1 column
  - Column 1: Opening balance for EPS (Employees' Pension Scheme)
- **Row 1**: Header row
- **Rows 2 onwards**: One row per employee

#### Sheet: WDL_EE
- **Columns**: 12 columns
  - Columns 1-12: Monthly withdrawals from Employee share (April to March)
- **Row 1**: Header row
- **Rows 2 onwards**: One row per employee

#### Sheet: WDL_ER
- **Columns**: 12 columns
  - Columns 1-12: Monthly withdrawals from Employer share (April to March)
- **Row 1**: Header row
- **Rows 2 onwards**: One row per employee

**Important**: All sheets must have the same number of rows (same number of employees).

### Output Files

The program generates two output files:

1. **output.csv**: Comma-separated values file
2. **Output.xlsx**: Formatted Excel file with:
   - Header row with colors and borders
   - All calculated columns
   - Center-aligned data

### Output Columns

| Column | Description |
|---------|-------------|
| A/C No. | Employee Account Number |
| NAME | Employee Name |
| OB(EE) | Opening Balance - Employee Share |
| OB(ER) | Opening Balance - Employer Share |
| INT(EE) | Interest Earned - Employee Share |
| INT(ER) | Interest Earned - Employer Share |
| CONT(EE) | Total Contribution - Employee Share (12%) |
| CONT(ER) | Total Contribution - Employer Share (3.67%) |
| WDL(EE) | Total Withdrawals - Employee Share |
| WDL(ER) | Total Withdrawals - Employer Share |
| CB(EE) | Closing Balance - Employee Share |
| CB(ER) | Closing Balance - Employer Share |
| OB(EPS) | Opening Balance - EPS |
| CONT(EPS) | Total Contribution - EPS (8.33%) |
| CB(EPS) | Closing Balance - EPS |

### Example Workflow

1. **Quick Test (Using Sample Data)**:
   ```bash
   cd epf_report_generator
   # Sample files already exist in InputFiles/
   python3 epf_calculator.py
   ```

2. **Prepare Your Own Input File**:
   ```bash
   cd epf_report_generator
   # Create or update InputFiles/Input.xlsx with all required sheets
   # Reference INPUT_FILE_TEMPLATE.md for detailed format
   ```

3. **Create New Sample**:
   ```bash
   cd epf_report_generator
   python3 create_sample_input.py
   # Generates Sample_Input.xlsx in InputFiles/
   ```

4. **Run Calculator**:
   ```bash
   cd epf_report_generator
   python3 epf_calculator.py
   ```

5. **Follow Prompts**:
   - Review instructions and press 'y' to continue
   - Enter annual interest rate (e.g., 8.5)
   - Wait for processing and validation

6. **Check Output**:
   - Review `output.csv` for data processing
   - Review `Output.xlsx` for formatted report

## Building Executable

### Windows Executable with PyInstaller

```bash
cd epf_report_generator
pip install pyinstaller
pyinstaller --onefile --windowed epf_calculator.py
```

Executable will be created in `dist/` directory.

### Windows Executable with cx_Freeze

```bash
cd epf_report_generator
pip install cx_Freeze
# Create setup.py with appropriate configuration
python setup.py build
```

## Testing

### Run Tests

```bash
pip install pytest
pytest
```

### Run Specific Test

```bash
pytest tests/test_calculator.py::test_calculate_contributions
```

### Run with Coverage

```bash
pytest --cov=. --cov-report=html
```

## Development

### Code Formatting

```bash
pip install black
black .
```

### Linting

```bash
pip install pylint
pylint epf_calculator.py
```

### Type Checking

```bash
pip install mypy
mypy epf_calculator.py
```

## EPF Calculation Logic

### Contribution Rates
- **Employee Contribution**: 12% of monthly wages
- **EPS Contribution**: 8.33% of monthly wages (part of employee contribution)
- **Employer Contribution**: 3.67% of monthly wages (remaining portion)

### Interest Calculation
```
Interest = round((Average Balance * Annual Rate) / 1200)
```

Where Average Balance is calculated by summing monthly balances and dividing by 12.

### Monthly Balance Tracking
For each month (starting from month 2):
```
Balance = Previous Balance + Current Month Contribution - Current Month Withdrawal
```

## Project Structure

```
epf_report_generator/
├── epf_calculator.py           # Main calculator program (consolidated version)
├── create_sample_input.py       # Script to generate sample input files
├── AGENTS.md                  # Development guidelines for AI assistants
├── README.md                  # This file
├── INPUT_FILE_TEMPLATE.md       # Detailed input file format specification
├── requirements.txt            # Python dependencies
├── setup.py                 # cx_Freeze configuration for Windows executable
├── .gitignore              # Version control exclusions
├── InputFiles/              # Input directory with sample files
│   ├── Input.xlsx           # Sample input ready to use
│   └── Sample_Input.xlsx    # Generated sample template
├── output.csv               # Generated CSV output (after running)
├── Output.xlsx              # Generated Excel output (after running)
└── tests/                  # Test directory
    └── test_calculator.py   # Test suite
```

## Error Handling

The program includes comprehensive error handling:

- **FileLoadError**: Input file cannot be loaded
- **SheetNotFoundError**: Required sheet is missing
- **DataValidationError**: Sheet dimensions don't match requirements
- **EPFCalculatorError**: Base exception for all calculator errors

All errors are displayed with clear messages in red color on the console.

## Notes

- This tool is designed for Windows OS due to console commands (cls, color, title)
- For Linux/Mac, consider modifying or removing `subprocess.call()` commands
- Existing output files (`output.csv`, `Output.xlsx`) will be overwritten
- All calculations follow EPF organization standards

## Version History

This is the **final consolidated version** combining features from:
- CSV-based calculators (Calc0.0 through Calc1.4)
- Excel-based calculators (ModifiedCalc variants)
- Enhanced with modern Python practices, type hints, and error handling

## 📄 License

This is an internal tool for EPF annual account slip generation and is available under the [MIT License](LICENSE).

## 👨‍💻 Author

**Jogesh Kumar Ghadai**
- Email: jogesh6895@gmail.com
- GitHub: [@Jogesh6895](https://github.com/Jogesh6895)

---
