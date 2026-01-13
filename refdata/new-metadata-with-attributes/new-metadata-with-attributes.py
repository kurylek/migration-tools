import os
import pandas as pd


def generate_refdata_metadata_xml(changeset_id, attribute_key, attribute_value_type, description):
    xml_template = f"""\n
    <changeSet author="ui-refdata" id="{changeset_id}">
    
        <insert tableName="refdata_metadata">
            <column name="attribute_key" value="{attribute_key}"/>
            <column name="attribute_value_type" value="{attribute_value_type}"/>
            <column name="description" value="{description}"/>
            <column name="owner_cl" value="CUSTOM"/>
        </insert>"""
    return xml_template


def generate_codelist_refdata_metadata_xml(codelist_key, attribute_key):
    xml_template = f"""\n
        <insert tableName="codelist_refdata_metadata">
            <column name="codelist_id" valueComputed="(SELECT id FROM codelist WHERE codelist_key = '{codelist_key}')"/>
            <column name="refdata_metadata_id" valueComputed="(SELECT id FROM refdata_metadata WHERE attribute_key = '{attribute_key}')"/>
            <column name="owner_cl" value="CUSTOM"/>
        </insert>"""
    return xml_template


def generate_refdata_attribute_xml(attribute_value, item_code, codelist_key, attribute_key):
    xml_template = f"""\n
        <insert tableName="refdata_attribute">
            <column name="attribute_value" value="{attribute_value}"/>
            <column name="item_id" valueComputed="(
                SELECT ci.id
                FROM codelist_item ci
                JOIN codelist c ON ci.codelist_id = c.id
                WHERE ci.item_code = '{item_code}'
                    AND c.codelist_key = '{codelist_key}'
            )"/>
            <column name="metadata_id" valueNumeric="(SELECT id FROM refdata_metadata WHERE attribute_key = '{attribute_key}')"/>
            <column name="owner_cl" value="CUSTOM"/>
        </insert>"""
    return xml_template


def read_excel_and_generate_xml(file_path):
    df = pd.read_excel(file_path, header=None)

    rule_data = df.iloc[1]
    migration_number = rule_data.iloc[0]
    attribute_key = rule_data.iloc[1]
    attribute_value_type = rule_data.iloc[2]
    description = rule_data.iloc[3]

    changeset_id = f"{migration_number}-EXT-INSERT-{attribute_key}-metadata"

    xml_output = generate_refdata_metadata_xml(changeset_id, attribute_key, attribute_value_type, description)

    codelist_data = df.iloc[3]

    for col_idx in range(1, len(codelist_data)):
        codelist_key = codelist_data.iloc[col_idx]

        if pd.notna(codelist_key) and str(codelist_key).strip():
            xml_output += generate_codelist_refdata_metadata_xml(codelist_key, attribute_key)

    for row_idx in range(6, len(df)):
        row_data = df.iloc[row_idx]

        if pd.notna(row_data.iloc[0]):
            codelist_key = row_data.iloc[0]
            item_code = row_data.iloc[1]
            attribute_value = row_data.iloc[2]

            xml_output += generate_refdata_attribute_xml(attribute_value, item_code, codelist_key, attribute_key)

    xml_output += "\n    </changeSet>"

    return xml_output, changeset_id


def write_to_file(content, file_path):
    header = '''<databaseChangeLog xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                   xmlns="http://www.liquibase.org/xml/ns/dbchangelog"
                   xsi:schemaLocation="http://www.liquibase.org/xml/ns/dbchangelog http://www.liquibase.org/xml/ns/dbchangelog/dbchangelog-4.6.xsd">'''

    footer = "\n</databaseChangeLog>\n"

    full_content = header + content + footer

    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(full_content)


script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, 'new-metadata-with-attributes.xlsm')
changesets, changeset_id = read_excel_and_generate_xml(file_path)
output_file_path = f'{changeset_id}.xml'
write_to_file(changesets, output_file_path)
print(f"XML output has been written to {output_file_path}")
