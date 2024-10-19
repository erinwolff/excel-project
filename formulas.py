import openpyxl
import json


def extract_formulas(file_path):
    # Load the workbook and select all the sheets
    workbook = openpyxl.load_workbook(file_path, data_only=False)

    # Dictionary to store formulas from all sheets
    formulas = {}

    # Iterate through each sheet in the workbook
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet_formulas = {}

        # Iterate through each cell in the sheet
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == "f":  # Check if the cell contains a formula
                    sheet_formulas[cell.coordinate] = cell.value

        # Add formulas from the sheet to the overall dictionary
        if sheet_formulas:
            formulas[sheet_name] = sheet_formulas

    return formulas


def save_formulas_to_json_file(formulas, output_file):
    with open(output_file, "w") as f:
        json.dump(formulas, f, indent=4)


# Usage
file_path = "BM055unprotected.xlsx"
formulas = extract_formulas(file_path)
save_formulas_to_json_file(formulas, "extracted_formulas.json")
