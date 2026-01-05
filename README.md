# XML Generators

## Description

These scripts processes Excel files and generates corresponding XML files for Liquibase migration.

## Requirements

Install Python dependencies:

```bash
pip install -r {PATH_TO_SCRIPT/}requirements.txt
```

or

```bash
py -m pip install -r {PATH_TO_SCRIPT/}requirements.txt
```

## Usage

1. Prepare your data in Excel file
2. Run the script:

```bash
py {SCRIPT_PATH}.py
```

3. The script will generate an XML file

## Output

The script generates a Liquibase XML changeset
