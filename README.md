# HEI Dataset and Queries Exporter

This project automates the process of exporting dataset schemas and SQL scripts from CSV files into a structured Excel workbook. The generated Excel workbooks include:

- An index sheet with links to dataset and script sheets.
- Separate sheets for each dataset schema.
- Separate sheets for each SQL script.
- A refresh instructions sheet detailing how to update the data.

## Prerequisites

- Python 3.x
- Virtual environment (recommended)
- Required Python packages: `pandas`, `openpyxl`

## Setup

1. **Clone the Repository:**

   ```sh
   git clone https://github.com/EddieDavison92/hei-dataset-and-queries-exporter.git
   cd hei-dataset-and-queries-exporter
   ```
2. **Create a Virtual Environment:**

   ```sh
   python -m venv venv
   ```
3. **Activate the Virtual Environment:**

   - On Windows:
     ```sh
     venv\Scripts\activate
     ```
   - On macOS and Linux:
     ```sh
     source venv/bin/activate
     ```
4. **Install Required Packages:**

   ```sh
   pip install -r requirements.txt
   ```

## Input Files

The following CSV files are required in the `input` directory:

- `HEI_V_CATALOG_LTCLCS.csv` (or other project-specific V_CATALOG file)
- `HEI_LTCLCS_SCRIPTS.csv` (or other project-specific scripts file)
- `HEI_V_CATALOG.csv` (complete Vertica catalog)
- `HEI_ALL_SCRIPTS.csv` (all scripts file)

Place Tableau workbooks in `.twb` format in the `input/tableau` directory.

## Usage

### Generating Excel Workbook

1. Modify the constants in any of the `create_catalog` functions to load the correct files and update titles and text as appropriate.
2. Run the corresponding script to generate an Excel file. For example:

   ```sh
   python create_ltclcs_catalog.py
   ```

### Refresh Instructions

The generated Excel workbook includes a "Refresh Instructions" sheet with detailed steps on how to update the data. Follow these instructions to refresh the datasets and SQL scripts.

## File Descriptions

### `create_<project>_catalog.py`

Scripts to generate specific project catalog Excel workbooks.

### `generate_excel.py`

Main execution script to read data, apply styles, and export to Excel.

### `modules/excel_helpers.py`

Contains helper functions for formatting and populating the Excel workbook.

### `modules/readers.py`

Contains functions to read CSV files and parse data.

### `write_instructions.py`

Adds a refresh instructions sheet to the Excel workbook.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

For any questions or issues, please refer to the repository or contact the project maintainers.
