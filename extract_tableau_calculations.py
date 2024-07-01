import os
import csv
import re
from tableaudocumentapi import Workbook

# Define the input and output paths
input_folder = 'input/tableau'
output_file = 'output/tableau.csv'

# Ensure the output directory exists
os.makedirs(os.path.dirname(output_file), exist_ok=True)

# Function to replace calculation IDs and sqlproxy references with field names
def replace_calc_ids_with_names(calculation, field_name_map, proxy_name_map):
    # Use regular expressions to find calculation IDs and sqlproxy references in the calculation string
    calc_pattern = re.compile(r'\[(Calculation_\d+)\]')
    sqlproxy_pattern = re.compile(r'\[sqlproxy\.(\w+)\]\.\[([^\]]+)\]')

    # Replace each calculation ID with the corresponding field name
    for match in calc_pattern.findall(calculation):
        field_name = field_name_map.get(match, match)
        calculation = calculation.replace(f'[{match}]', f'[{field_name}]')

    # Replace each sqlproxy reference with the corresponding data source and field name
    for match in sqlproxy_pattern.findall(calculation):
        proxy_key, proxy_field = match
        full_match = f'[sqlproxy.{proxy_key}].[{proxy_field}]'
        if proxy_key in proxy_name_map:
            replacement = f'[{proxy_name_map[proxy_key]}].[{proxy_field}]'
            calculation = calculation.replace(full_match, replacement)
    
    return calculation

# List to hold the calculated field details
calculated_fields = []

# Iterate over all .twb files in the input folder
for filename in os.listdir(input_folder):
    if filename.endswith('.twb'):
        filepath = os.path.join(input_folder, filename)
        workbook = Workbook(filepath)

        # Extract calculated fields from the workbook
        for datasource in workbook.datasources:
            datasource_name = datasource.caption if datasource.caption else datasource.name
            if datasource_name.lower() == 'parameters':
                continue

            # Build a mapping of calculation IDs to field names
            field_name_map = {}
            proxy_name_map = {}
            for field in datasource.fields.values():
                if field.calculation and not field.calculation.isspace():
                    # Create a mapping from calculation ID to field name
                    calculation_id = re.search(r'\[Calculation_(\d+)\]', field.calculation)
                    if calculation_id:
                        field_name_map[f'Calculation_{calculation_id.group(1)}'] = field.name
                    # Create a mapping for sqlproxy references
                    proxy_pattern = re.compile(r'\[sqlproxy\.(\w+)\]\.\[([^\]]+)\]')
                    for proxy_key, proxy_field in proxy_pattern.findall(field.calculation):
                        proxy_name_map[proxy_key] = datasource_name

            for field in datasource.fields.values():
                if field.calculation and not field.calculation.isspace():
                    # Replace calculation IDs and sqlproxy references with field names in the calculation string
                    calculation_with_names = replace_calc_ids_with_names(
                        field.calculation.strip(), field_name_map, proxy_name_map)
                    calculated_fields.append({
                        'Workbook': filename,
                        'Data Source': datasource_name,
                        'Field Name': field.name,
                        'Calculation': calculation_with_names.replace('\n', ' '),
                        'Data Type': field.datatype
                    })

# Sort the calculated fields first by workbook, then by data source, then by field name
calculated_fields.sort(key=lambda x: (x['Workbook'], x['Data Source'], x['Field Name']))

# Define the CSV fieldnames
fieldnames = ['Workbook', 'Data Source', 'Field Name', 'Calculation', 'Data Type']

# Write the calculated fields to the CSV file
with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for field in calculated_fields:
        writer.writerow(field)

print(f'Calculated fields have been extracted to {output_file}')