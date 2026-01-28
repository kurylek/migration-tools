# Form Rule XML Generator

## Description

This script processes Excel files containing query definitions and generates corresponding XML files for Liquibase migration.

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

1. Prepare your data in Excel file named `new-query.xlsm`
2. Run the script:
   - ```bash
       py new-query.py
       ```
   - Macro in Excel file

3. The script will generate an XML file

## Output

The script generates a Liquibase XML changeset
