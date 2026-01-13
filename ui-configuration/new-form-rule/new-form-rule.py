import os
import pandas as pd


def generate_form_rule_xml(changeset_id, rule_name, category_cl, use_cl, comment, description_en, description_lt):
    xml_template = f"""\n
    <changeSet author="ui-configuration" id="{changeset_id}">
        <insert tableName="c_form_rule">
            <column name="form_rule_cd" value="{rule_name}"/>
            <column name="form_rule_category_cl" value="{category_cl}"/>
            <column name="form_rule_use_cl" value="{use_cl}"/>
            <column name="comment" value="{comment}"/>
            <column name="owner_cl" value="CUSTOM"/>
        </insert>

        <insert tableName="c_form_rule_l">
            <column name="form_rule_cd" value="{rule_name}"/>
            <column name="language_cl" value="EN"/>
            <column name="description" value="{description_en}"/>
        </insert>

        <insert tableName="c_form_rule_l">
            <column name="form_rule_cd" value="{rule_name}"/>
            <column name="language_cl" value="LT"/>
            <column name="description" value="{description_lt}"/>
        </insert>"""
    return xml_template

def generate_exception_type_xml(rule_name, exception_type_cd):
    xml_template = f"""\n
        <insert tableName="c_form_rule_exception_type">
            <column name="form_rule_cd" value="{rule_name}"/>
            <column name="exception_type_cd" value="{exception_type_cd}"/>
        </insert>"""
    return xml_template

def generate_parameter_xml(rule_name, parameter_name, data_type, required, parameter_type, parameter_comment):
    if pd.isna(parameter_comment) or parameter_comment == "":
        param_comment = '<column name="comment"/>'
    else:
        param_comment = f'<column name="comment" value="{parameter_comment}"/>'

    xml_template = f"""\n
        <insert tableName="c_form_rule_parameter">
            <column name="form_rule_cd" value="{rule_name}"/>
            <column name="parameter_name" value="{parameter_name}"/>
            <column name="form_rule_parameter_data_type_cl" value="{data_type}"/>
            <column name="required_sw" value="{required}"/>
            <column name="form_rule_parameter_type_cl" value="{parameter_type}"/>
            {param_comment}
        </insert>"""
    return xml_template


def read_excel_and_generate_xml(file_path):
    df = pd.read_excel(file_path, header=None)

    rule_data = df.iloc[1]
    migration_number = rule_data.iloc[0]
    rule_name = rule_data.iloc[1]
    category_cl = rule_data.iloc[2]
    use_cl = rule_data.iloc[3]
    comment = rule_data.iloc[4]
    description_en = rule_data.iloc[5]
    description_lt = rule_data.iloc[6]

    changeset_id = f"{migration_number}-EXT-INSERT-{rule_name}-rule"

    xml_output = generate_form_rule_xml(changeset_id, rule_name, category_cl, use_cl, comment, description_en, description_lt)

    exception_data = df.iloc[3]

    for col_idx in range(1, len(exception_data)):
        exception_type_cd = exception_data.iloc[col_idx]

        if pd.notna(exception_type_cd) and str(exception_type_cd).strip():
            xml_output += generate_exception_type_xml(rule_name, exception_type_cd)

    for row_idx in range(6, len(df)):
        row_data = df.iloc[row_idx]

        if pd.notna(row_data.iloc[0]):
            parameter_name = row_data.iloc[0]
            data_type = row_data.iloc[1]
            required = row_data.iloc[2]
            parameter_type = row_data.iloc[3]
            parameter_comment = row_data.iloc[4]

            xml_output += generate_parameter_xml(rule_name, parameter_name, data_type, required, parameter_type, parameter_comment)

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
file_path = os.path.join(script_dir, 'new-form-rule.xlsm')
changesets, changeset_id = read_excel_and_generate_xml(file_path)
output_file_path = f'{changeset_id}.xml'
write_to_file(changesets, output_file_path)
print(f"XML output has been written to {output_file_path}")
