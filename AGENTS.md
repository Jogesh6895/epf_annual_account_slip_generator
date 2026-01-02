# AGENTS.md

This guide is for agentic coding assistants working on TotalCalculator project.

## Project Overview

This is an EPF (Employees' Provident Fund) annual account slip calculator with Excel input/output support.

**Latest Version**: `epf_report_generator/epf_calculator.py` - Consolidated version combining all features from:
- CSV-based calculators (Calc0.0 through Calc1.4)
- Excel-based calculators (ModifiedCalc variants)
- Enhanced with type hints, custom exceptions, and comprehensive validation

### Quick Start

```bash
cd epf_report_generator
# Sample input files already provided
python3 epf_calculator.py
```

## Build and Run Commands

### Running the Calculator

```bash
cd epf_report_generator
python3 epf_calculator.py
```

### Generating Sample Input

```bash
cd epf_report_generator
python3 create_sample_input.py
# Creates InputFiles/Sample_Input.xlsx with 5 dummy employees
```

### Running with Sample Data

```bash
cd epf_report_generator
# Sample files (Input.xlsx and Sample_Input.xlsx) already provided
python3 epf_calculator.py
```

Note: Use `python3` on Linux/Mac systems. Use `python` on Windows.

### Building Windows Executables

```bash
# Build CSV calculator with cx_Freeze
cd Calculator/Calculator/Calc1.4/Source
python setup.py build

# Build CSV calculator with PyInstaller
cd Calculator/Calculator/Calc1.4/pyinstaller-1
pyinstaller Calculator.py

# Build Excel calculator
cd FromDesktop/ModifiedCalc
# Add setup.py if creating executable
```

### Testing Commands (Recommended Setup)

```bash
# Install pytest
pip install pytest

# Run all tests
pytest

# Run a single test file
pytest tests/test_calculator.py

# Run a specific test function
pytest tests/test_calculator.py::test_calculate_interest

# Run with verbose output
pytest -v

# Run with coverage
pytest --cov=. --cov-report=html
```

### Linting and Formatting (Recommended Setup)

```bash
# Install tools
pip install black pylint mypy

# Format code with Black
black .

# Check code quality
pylint FromDesktop/ModifiedCalc/ModifiedCalculator.py
pylint Calculator/Calculator/Calc1.4/Source/Calculator.py

# Type checking
mypy FromDesktop/ModifiedCalc/ModifiedCalculator.py
mypy Calculator/Calculator/Calc1.4/Source/Calculator.py
```

## Code Style Guidelines

### Naming Conventions
- **Functions and variables:** `snake_case` - e.g., `open_input_excel_file()`, `wage_sheet_data`
- **Constants:** `UPPER_CASE` - e.g., `EMPLOYEE_CONTRIBUTION_RATE`, `EXPECTED_COLUMNS`
- **Classes:** `PascalCase` - e.g., `WageCalculator`, `DataValidator`
- **Private functions:** Prefix with `_` - e.g., `_validate_sheet_data()`

### Import Organization
```python
# 1. Standard library imports first
import csv
import subprocess
import time

# 2. Third-party imports second
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, colors

# 3. Local imports third (if any)
from .utils import validate_data
```

**Guidelines:**
- Wildcard imports acceptable for style modules only: `from openpyxl.styles import *`
- Group imports by type with blank lines between groups
- Sort imports alphabetically within each group

### Type Hints
Add type hints for better code clarity:
```python
from typing import List, Dict, Union, Optional

def open_input_excel_file(path: str) -> Union[openpyxl.Workbook, str]:
    try:
        return openpyxl.load_workbook(path)
    except Exception:
        return 'ERROR'

def calculate_interest(opening_balance: float, rate: float) -> int:
    return round((opening_balance * rate) / 1200)
```

### Error Handling
**Current pattern (observed):** Return string 'ERROR' on exceptions
```python
def open_input_excel_file(path: str) -> Union[openpyxl.Workbook, str]:
    try:
        return openpyxl.load_workbook(path)
    except Exception:
        return 'ERROR'
```

**Recommended improvement:** Use specific exceptions
```python
class WorkbookLoadError(Exception):
    pass

class SheetNotFoundError(Exception):
    pass

def open_input_excel_file(path: str) -> openpyxl.Workbook:
    try:
        return openpyxl.load_workbook(path)
    except FileNotFoundError:
        raise WorkbookLoadError(f"File not found: {path}")
    except Exception as e:
        raise WorkbookLoadError(f"Failed to load workbook: {e}")
```

### Formatting
- Use **4 spaces** for indentation (no tabs)
- Maximum line length: 100 characters (Black default is 88)
- Add blank lines between functions (2 blank lines for top-level functions)
- One space around operators: `tee = round(float(i) * 0.12)`

### Function Design
**Use early returns and guard clauses instead of deep nesting:**

Current pattern (deep nesting):
```python
if input_workbook != 'ERROR':
    if wage_sheet != 'ERROR':
        if ob_ee_sheet != 'ERROR':
            # deep nesting...
```

Preferred pattern (early returns):
```python
if input_workbook == 'ERROR':
    print("ERROR! Input.xlsx file not found")
    return

if wage_sheet == 'ERROR':
    print("ERROR! 'Wages' Sheet not found")
    return

if ob_ee_sheet == 'ERROR':
    print("ERROR! 'OB_EE' Sheet not found")
    return

# Main logic here
```

### File I/O
Always use context managers for file operations:
```python
# CSV reading
with open('input.csv', 'r') as csvfile:
    reader = csv.reader(csvfile, delimiter=',')
    for row in reader:
        process_row(row)

# CSV writing
with open('output.csv', 'w', newline='') as outputfile:
    writer = csv.writer(outputfile, delimiter=',')
    writer.writerow(header_row)
```

### Constants
Define constants at module level for magic numbers:
```python
EMPLOYEE_CONTRIBUTION_RATE = 0.12
EPS_CONTRIBUTION_RATE = 0.0833
EMPLOYER_CONTRIBUTION_RATE = 0.0367
INTEREST_DIVISOR = 1200
EXPECTED_WAGES_COLUMNS = 14
EXPECTED_WDL_COLUMNS = 12
EXPECTED_OB_COLUMNS = 1
```

## Latest Versions

### CSV-based Calculator: Calculator/Calculator/Calc1.4/Source/Calculator.py
**Features:**
- Reads CSV files from Input/ directory
- Calculates EPF contributions, interest, and closing balances
- Outputs to output.csv
- Windows console formatting with subprocess commands
- Time tracking for performance monitoring
- Packaging support with cx_Freeze and PyInstaller

**Input Files (in Input/ directory):**
- `input.csv` - Employee names, account numbers, and monthly wages
- `OB(EE).csv` - Opening balance for employee share
- `OB(ER).csv` - Opening balance for employer share
- `OB(EPS).csv` - Opening balance for EPS
- `WDL(EE).csv` - Monthly withdrawals from employee share
- `WDL(ER).csv` - Monthly withdrawals from employer share

**Output File:**
- `output.csv` - Complete annual account slip with all calculations

### Excel-based Calculator: FromDesktop/ModifiedCalc/ModifiedCalculator.py
**Features:**
- Reads Excel file from InputFiles/Input.xlsx
- Comprehensive validation (column and row counts)
- Data fetching with null value handling
- Startup confirmation dialog
- Screen clearing and formatted output
- Detailed error messages for missing sheets/files
- Modular structure with helper functions

**Input File:**
- `InputFiles/Input.xlsx` with required sheets:
  - `Wages` - Monthly wage data (14 columns)
  - `OB_EE` - Opening balance for employee share (1 column)
  - `OB_ER` - Opening balance for employer share (1 column)
  - `OB_EPS` - Opening balance for EPS (1 column)
  - `WDL_EE` - Monthly withdrawals from employee share (12 columns)
  - `WDL_ER` - Monthly withdrawals from employer share (12 columns)

**Validation Rules:**
- All sheets must have matching row counts
- Wages sheet: 14 columns
- OB sheets: 1 column each
- WDL sheets: 12 columns each

## EPF Calculation Logic

Both calculators use the same EPF contribution rates and formulas:

### Contribution Rates
- Employee contribution: 12% of wages
- EPS contribution: 8.33% of wages
- Employer contribution: 3.67% of wages (12% - 8.33%)

### Interest Calculation
```python
interest = round((average_balance * rate) / 1200)
```
Where rate is the annual interest rate entered by user.

### Monthly Balance Tracking
For each month (excluding first):
```python
balance = previous_balance + contribution - withdrawal
```

## Version Structure

### Reference Directories (For Historical Reference)
- `Calculator/` - CSV-based calculator versions (Calc0.0 through Calc1.4)
- `ModifiedCalc/` - Early Excel-based calculator versions
- `FromDesktop/` - Archive/backup of older versions
- `Previous_CSV_Calc/` - Backup of CSV calculators

**Note**: These directories are for reference only. Use `epf_report_generator/epf_calculator.py` for all new development.

## Development Workflow

1. **Make changes** in `epf_report_generator/epf_calculator.py`
2. **Test changes** with sample data
   ```bash
   python3 epf_calculator.py
   ```
3. **Format code** with `black .`
4. **Run linting** with `pylint`
5. **Run tests** with `pytest`
6. **Build executable** (if needed) with `python3 setup.py build`

## Important Notes

- **Windows-specific commands:** `subprocess.call()` is used for Windows console commands (cls, color, title). These will not work on Linux/Mac.
- **Cross-platform consideration:** Consider adding platform detection for better portability.
- **Error handling pattern:** Current code uses custom exceptions with proper error messages.
- **Data validation:** Comprehensive validation for sheet dimensions, column counts, and row counts.
- **Sample files:** Pre-generated sample input files are provided in `InputFiles/` directory.
- **Testing**: Sample data includes 5 employees with various wage ranges and withdrawal scenarios.

## Available Reference Versions

The following directories contain older versions for reference only (DO NOT MODIFY):

**CSV Calculators**:
- `Calculator/Calculator/Calc0.0/` (V-1.0 through V-1.5)
- `Calculator/Calculator/Calc1.0/`, `Calc1.1/`, `Calc1.2/`
- `Calculator/Calculator/Calc1.3/`, `Calc1.4/Source/`

**Excel Calculators**:
- `ModifiedCalc/ModifiedCalculator.py` (125 lines)
- `Calculator/ModifiedCalc/ModifiedCalculator.py` (143 lines)
- `FromDesktop/ModifiedCalc/ModifiedCalculator.py` (303 lines)

**Archives**:
- `FromDesktop/` - Archive of various versions
- `Previous_CSV_Calc/` - Backup directory
