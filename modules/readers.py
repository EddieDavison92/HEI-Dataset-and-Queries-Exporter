import pandas as pd

def read_csv_schema(file_path):
    df = pd.read_csv(file_path)
    tables = {}
    for _, row in df.iterrows():
        table_schema = row['table_schema']
        table_name = row['table_name']
        column_name = row['column_name']
        data_type = row['data_type']
        ordinal_position = row['ordinal_position']
        if table_name not in tables:
            tables[table_name] = {'schema': table_schema, 'columns': []}
        tables[table_name]['columns'].append((ordinal_position, column_name, data_type))
    return tables

def read_csv_scripts(file_path):
    return pd.read_csv(file_path)

def read_csv_tableau(file_path, workbook_name):
    df = pd.read_csv(file_path)
    return df[df['Workbook'] == workbook_name]

def read_additional_v_catalog(file_path):
    return pd.read_csv(file_path)

def read_all_scripts(file_path):
    return pd.read_csv(file_path)
