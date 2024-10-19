import openpyxl


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


# Usage
file_path = "your_excel_file.xlsx"
formulas = extract_formulas(file_path)

# Print the extracted formulas
for sheet, sheet_formulas in formulas.items():
    print(f"Formulas in sheet '{sheet}':")
    for cell, formula in sheet_formulas.items():
        print(f"Cell {cell}: {formula}")
