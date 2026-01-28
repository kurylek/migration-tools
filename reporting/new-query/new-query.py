import os
import pandas as pd


def generate_query_xml(changeset_id, query_cd, query_grouping_cl, description_en, description_lt):
    xml_template = f"""\n
    <changeSet author="reporting" id="{changeset_id}">
        <insert tableName="c_query">
            <column name="query_cd" value="{query_cd}"/>
            <column name="query_grouping_cl" value="{query_grouping_cl}"/>
            <column name="owner_cl" value="CUSTOM"/>
            <column name="query_text"
                    valueClobFile="../native-sql/postgres/queries/{query_cd}/query-text.sql"/>
            <column name="count_query_text"
                    valueClobFile="../native-sql/postgres/queries/{query_cd}/count-query-text.sql"/>
        </insert>

        <insert tableName="c_query_l">
            <column name="query_cd" value="{query_cd}"/>
            <column name="language_cl" value="EN"/>
            <column name="description"
                    value="{description_en}"/>
        </insert>

        <insert tableName="c_query_l">
            <column name="query_cd" value="{query_cd}"/>
            <column name="language_cl" value="LT"/>
            <column name="description"
                    value="{description_lt}"/>
        </insert>
    </changeSet>"""
    return xml_template


def generate_securable_objects_xml(changeset_id, query_cd, role_cd):
    xml_template = f"""\n
    <changeSet author="authorization-management" id="{changeset_id}">
        <insert tableName="c_securable_object_operation">
            <column name="securable_object_key" value="ENTITY/query-management/queries/"/>
            <column name="securable_object_operation_key" value="{query_cd}/execute/POST"/>
            <column name="owner_cl" value="CUSTOM"/>
        </insert>

        <insert tableName="c_role_securable_object_operation">
            <column name="role_cd" value="{role_cd}"/>
            <column name="securable_object_key" value="ENTITY/query-management/queries/"/>
            <column name="securable_object_operation_key" value="{query_cd}/execute/POST"/>
        </insert>
    </changeSet>"""
    return xml_template


def read_excel_and_generate_xml(file_path):
    df = pd.read_excel(file_path, header=None)

    query_data = df.iloc[1]
    query_migration_number = query_data.iloc[0]
    query_cd = query_data.iloc[1]
    query_grouping_cl = query_data.iloc[2]
    description_en = query_data.iloc[3]
    description_lt = query_data.iloc[4]

    securable_data = df.iloc[4]
    securable_migration_number = securable_data.iloc[0]
    role_cd = securable_data.iloc[1]

    reporting_changeset_id = f"{query_migration_number}-EXT-INSERT-{query_cd}-QUERY"
    auth_changeset_id = f"{securable_migration_number}-EXT-INSERT-{query_cd}-QUERY-SECURABLE_OBJECTS"

    query_xml_output = generate_query_xml(reporting_changeset_id, query_cd, query_grouping_cl, description_en, description_lt)
    auth_xml_output = generate_securable_objects_xml(auth_changeset_id, query_cd, role_cd)

    return query_xml_output, reporting_changeset_id, auth_xml_output, auth_changeset_id, query_cd


def write_to_file(content, file_path):
    header = '''<databaseChangeLog xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                   xmlns="http://www.liquibase.org/xml/ns/dbchangelog"
                   xsi:schemaLocation="http://www.liquibase.org/xml/ns/dbchangelog http://www.liquibase.org/xml/ns/dbchangelog/dbchangelog-4.6.xsd">'''

    footer = "\n</databaseChangeLog>\n"

    full_content = header + content + footer

    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(full_content)

def create_placeholder_file(file_path):
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    with open(file_path, 'w') as file:
        pass


script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, 'new-query.xlsm')
reporting_changeset, reporting_changeset_id, auth_changeset, auth_changeset_id, query_cd = read_excel_and_generate_xml(file_path)

reporting_output_file_path = f'out/reporting/{reporting_changeset_id}.xml'
auth_output_file_path = f'out/authorization-management/{auth_changeset_id}.xml'

os.makedirs(os.path.dirname(reporting_output_file_path), exist_ok=True)
os.makedirs(os.path.dirname(auth_output_file_path), exist_ok=True)

write_to_file(reporting_changeset, reporting_output_file_path)
write_to_file(auth_changeset, auth_output_file_path)

create_placeholder_file(f'out/reporting/{query_cd}/query-text.sql')
create_placeholder_file(f'out/reporting/{query_cd}/count-query-text.sql')

print(f"XML output has been written to out directory")
