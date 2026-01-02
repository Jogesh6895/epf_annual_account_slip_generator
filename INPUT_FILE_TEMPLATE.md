# Input File Template

This document describes the structure of `InputFiles/Input.xlsx` required for the EPF Report Generator.

## File Location
- Path: `InputFiles/Input.xlsx`
- Must be created by user before running the calculator

## Required Sheets

### Sheet 1: Wages
This sheet contains employee information and monthly wages.

| Column | Letter | Header | Description |
|---------|----------|----------|-------------|
| 1 | A | A/C No. | Employee Account Number |
| 2 | B | NAME | Employee Name |
| 3 | C | Apr | Wages for April |
| 4 | D | May | Wages for May |
| 5 | E | Jun | Wages for June |
| 6 | F | Jul | Wages for July |
| 7 | G | Aug | Wages for August |
| 8 | H | Sep | Wages for September |
| 9 | I | Oct | Wages for October |
| 10 | J | Nov | Wages for November |
| 11 | K | Dec | Wages for December |
| 12 | L | Jan | Wages for January |
| 13 | M | Feb | Wages for February |
| 14 | N | Mar | Wages for March |

**Total Columns**: 14
**Example Row 1**:
```
A/C No.|NAME|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar
EPF001|John Doe|15000|15000|15000|...
```

### Sheet 2: OB_EE
Opening Balance - Employee Share

| Column | Letter | Header | Description |
|---------|----------|----------|-------------|
| 1 | A | OB(EE) | Opening Balance for Employee Share |

**Total Columns**: 1
**Example Data**:
```
OB(EE)
50000
45000
60000
```

### Sheet 3: OB_ER
Opening Balance - Employer Share

| Column | Letter | Header | Description |
|---------|----------|----------|-------------|
| 1 | A | OB(ER) | Opening Balance for Employer Share |

**Total Columns**: 1
**Example Data**:
```
OB(ER)
15000
13500
18000
```

### Sheet 4: OB_EPS
Opening Balance - EPS (Employees' Pension Scheme)

| Column | Letter | Header | Description |
|---------|----------|----------|-------------|
| 1 | A | OB(EPS) | Opening Balance for EPS |

**Total Columns**: 1
**Example Data**:
```
OB(EPS)
35000
31500
42000
```

### Sheet 5: WDL_EE
Withdrawals - Employee Share

| Column | Letter | Header | Description |
|---------|----------|----------|-------------|
| 1 | A | Apr | Withdrawal in April |
| 2 | B | May | Withdrawal in May |
| 3 | C | Jun | Withdrawal in June |
| 4 | D | Jul | Withdrawal in July |
| 5 | E | Aug | Withdrawal in August |
| 6 | F | Sep | Withdrawal in September |
| 7 | G | Oct | Withdrawal in October |
| 8 | H | Nov | Withdrawal in November |
| 9 | I | Dec | Withdrawal in December |
| 10 | J | Jan | Withdrawal in January |
| 11 | K | Feb | Withdrawal in February |
| 12 | L | Mar | Withdrawal in March |

**Total Columns**: 12
**Example Row**:
```
Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar
0|0|0|5000|0|0|0|0|0|0|0|0
```

### Sheet 6: WDL_ER
Withdrawals - Employer Share

| Column | Letter | Header | Description |
|---------|----------|----------|-------------|
| 1 | A | Apr | Withdrawal in April |
| 2 | B | May | Withdrawal in May |
| 3 | C | Jun | Withdrawal in June |
| 4 | D | Jul | Withdrawal in July |
| 5 | E | Aug | Withdrawal in August |
| 6 | F | Sep | Withdrawal in September |
| 7 | G | Oct | Withdrawal in October |
| 8 | H | Nov | Withdrawal in November |
| 9 | I | Dec | Withdrawal in December |
| 10 | J | Jan | Withdrawal in January |
| 11 | K | Feb | Withdrawal in February |
| 12 | L | Mar | Withdrawal in March |

**Total Columns**: 12
**Example Row**:
```
Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar
0|0|0|1500|0|0|0|0|0|0|0|0
```

## Important Rules

1. **Row Count Consistency**: All sheets must have the same number of rows (excluding header row)
2. **Header Row**: Row 1 in each sheet should contain column headers
3. **Data Rows**: Data starts from Row 2 in each sheet
4. **Employee Order**: Employees must be in the same order across all sheets
5. **Empty Cells**: Treat empty cells as zero (0)
6. **Sheet Names**: Sheet names must exactly match the names listed above (case-sensitive)

## Complete Example

### Wages Sheet (Row 1-3)
```
A/C No.|NAME|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar
EPF001|John Doe|15000|15000|15500|15000|15000|15000|16000|15000|15000|15500|15000|15000
EPF002|Jane Smith|18000|18000|18500|18000|18000|18000|19000|18000|18000|18500|18000|18000
```

### OB_EE Sheet (Row 1-3)
```
OB(EE)
50000
60000
```

### OB_ER Sheet (Row 1-3)
```
OB(ER)
15000
18000
```

### OB_EPS Sheet (Row 1-3)
```
OB(EPS)
35000
42000
```

### WDL_EE Sheet (Row 1-3)
```
Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar
0|0|0|5000|0|0|0|0|0|0|0|0
0|0|0|0|0|0|0|0|0|0|0|0
```

### WDL_ER Sheet (Row 1-3)
```
Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar
0|0|0|1500|0|0|0|0|0|0|0|0
0|0|0|0|0|0|0|0|0|0|0|0
```

## Creating Input File

1. Open Microsoft Excel or LibreOffice Calc
2. Create new workbook
3. Create 6 sheets with exact names: Wages, OB_EE, OB_ER, OB_EPS, WDL_EE, WDL_ER
4. Add headers as described above
5. Fill in employee data
6. Save as `InputFiles/Input.xlsx`
7. Run: `python epf_calculator.py`

## Validation

The program will validate:
- All sheets exist with correct names
- Column counts match requirements (Wages: 14, OB sheets: 1, WDL sheets: 12)
- Row counts match across all sheets
- Data is numeric where required

If validation fails, the program will show specific error message indicating which sheet and what issue was found.
