import os
import pandas as pd


def generate_form_type_rule_xml(changeset_id, form_name, rule_name, group, priority, comment, skip_on_error):
    xml_template = f"""\n
    <changeSet author="ui-configuration" id="{changeset_id}">
        <preConditions onFail="HALT" onFailMessage="Record with same (form_type_cd, form_type_rule_group_cd, form_rule_cd, comment) already exists in c_form_type_rule table">
            <sqlCheck expectedResult="0">
                SELECT COUNT(*)
                FROM c_form_type_rule
                WHERE form_type_cd = '{form_name}'
                  AND form_type_rule_group_cd = '{group}'
                  AND form_rule_cd = '{rule_name}'
                  AND comment = '{comment}'
            </sqlCheck>
        </preConditions>
    
        <insert tableName="c_form_type_rule">
            <column name="form_type_rule_id" valueComputed="uuid_generate_v4()"/>
            <column name="form_type_cd" value="{form_name}"/>
            <column name="form_type_rule_group_cd" value="{group}"/>
            <column name="form_rule_cd" value="{rule_name}"/>
            <column name="priority" valueNumeric="{priority}"/>
            <column name="skip_on_error_sw" value="{skip_on_error}"/>
            <column name="active_sw" value="Y"/>
            <column name="comment" value="{comment}"/>
        </insert>"""
    return xml_template


def generate_parameter_xml(rule_name, form_name, parameter_name, comment, parameter_path, parameter_value,
                           parameter_type):
    if parameter_type == "LITERAL":
        param_column = f'<column name="parameter_value" value="{parameter_value}"/>'
    elif parameter_type == "FROM_FORM":
        param_column = f'<column name="parameter_path" value="{parameter_path}"/>'
    else:
        raise ValueError(f"Unknown parameter_type: '{parameter_type}'. Allowed values: 'LITERAL', 'FROM_FORM'")

    xml_template = f"""\n
        <insert tableName="c_form_type_rule_parameter">
            <column name="form_type_rule_parameter_id" valueComputed="uuid_generate_v4()"/>
            <column name="form_type_rule_id" valueComputed="(
                    SELECT form_type_rule_id
                    FROM c_form_type_rule
                    WHERE form_rule_cd = '{rule_name}'
                    AND form_type_cd = '{form_name}'
                    AND comment = '{comment}'
                )"/>
            <column name="form_rule_cd" value="{rule_name}"/>
            <column name="parameter_name" value="{parameter_name}"/>
            <column name="form_type_rule_parameter_type_cl" value="{parameter_type}"/>
            {param_column}
        </insert>"""
    return xml_template


def read_excel_and_generate_xml(file_path):
    df = pd.read_excel(file_path, header=None)

    rule_data = df.iloc[1]
    migration_number = rule_data.iloc[0]
    form_name = rule_data.iloc[1]
    rule_name = rule_data.iloc[2]
    group = rule_data.iloc[3]
    priority = rule_data.iloc[4]
    skip_on_error = rule_data.iloc[5]
    comment = rule_data.iloc[6]

    changeset_id = f"{migration_number}-EXT-INSERT-{form_name}-{rule_name}-instance"

    xml_output = generate_form_type_rule_xml(changeset_id, form_name, rule_name, group, priority, comment,
                                             skip_on_error)

    for row_idx in range(4, len(df)):
        row_data = df.iloc[row_idx]

        if pd.notna(row_data.iloc[0]):
            parameter_name = row_data.iloc[0]
            parameter_type = row_data.iloc[1]
            parameter_path = row_data.iloc[2]
            parameter_value = row_data.iloc[3]

            xml_output += generate_parameter_xml(rule_name, form_name, parameter_name, comment, parameter_path,
                                                 parameter_value, parameter_type)

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
file_path = os.path.join(script_dir, 'new-form-rule-instance.xlsm')
changesets, changeset_id = read_excel_and_generate_xml(file_path)
output_file_path = f'{changeset_id}.xml'
write_to_file(changesets, output_file_path)
print(f"XML output has been written to {output_file_path}")
