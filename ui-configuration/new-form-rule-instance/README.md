# Form Rule XML Generator

## Description

This script processes Excel files containing form rule instance definitions and generates corresponding XML files for Liquibase migration.

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

1. Prepare your data in Excel file named `new-form-rule-instance.xlsx`
2. Run the script:

```bash
py new-form-rule-instance.py
```

3. The script will generate an XML file

## Output

The script generates a Liquibase XML changeset
