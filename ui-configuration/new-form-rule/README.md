# Form Rule XML Generator

## Description

This script processes Excel files containing form rule definitions and generates corresponding XML files for Liquibase migration.

## Requirements

Install Python dependencies:

```bash
pip install -r requirements.txt
```

or

```bash
py -m pip install -r requirements.txt
```

## Usage

1. Prepare your data in Excel file named `new-form-rule.xlsm`
2. Run the script:
    - ```bash
        py new_form_rule.py
        ```
    - Macro in Excel file

3. The script will generate an XML file

## Output

The script generates a Liquibase XML changeset containing:

- **Form Rule Insert**: Main rule definition with metadata
- **Multilingual Descriptions**: English and Lithuanian descriptions
- **Exception Types**: Associated exception types (if any)
- **Parameters**: Rule parameters with data types and constraints

## Notes

- The script supports variable number of exception types (columns 1+ in row 3)
- Parameters start from row 6 and continue until the end of the file
