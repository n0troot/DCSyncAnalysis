# DCSyncAnalysis

A powerful password analysis tool that processes NTLM hash dumps and cracked passwords to generate comprehensive Excel reports with visualizations.

## Features

- **Password Analysis**: Analyze cracked passwords from NTLM hash dumps
- **Success Rate Tracking**: Calculate and visualize password cracking success rates
- **Common Password Detection**: Identify and chart the most commonly used passwords
- **Pattern Analysis**: Analyze password complexity patterns (uppercase, lowercase, numbers, special chars)
- **Length Distribution**: Visualize the distribution of password lengths
- **Season/Year Detection**: Identify passwords containing seasons and years
- **Rich Visualizations**: Generate bar charts, pie charts, and formatted tables
- **Excel Report**: Output all analysis in a professional Excel workbook

## Requirements

- Python 3.6+
- Required packages:
  - pandas
  - openpyxl

## Installation

```bash
pip install pandas openpyxl
```

## Usage

```bash
python DCSyncAnalysis.py <hashes_file> <cracked_file> <output_file> <company_name>
```

### Parameters

- `hashes_file`: File containing NTLM hashes (secretsdump.py format)
- `cracked_file`: File containing cracked passwords (hash:password format)
- `output_file`: Excel output file for the report (.xlsx)

## Example

```bash
python DCSyncAnalysis.py domain_hashes.txt cracked_passwords.txt password_analysis.xlsx COMAPANY_NAME
```

## Output

The tool generates an Excel workbook with multiple sheets:

1. **Summary**: Overall statistics and cracking success rate
2. **Common Passwords**: Most frequently used passwords with charts
3. **Password Patterns**: Analysis of character types and complexity
4. **Length Analysis**: Distribution of password lengths
5. **Time Patterns**: Analysis of season/year patterns in passwords
