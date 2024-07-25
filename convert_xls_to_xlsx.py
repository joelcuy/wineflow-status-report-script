import os
import xlrd
import openpyxl
import pandas as pd


def convert_xls_to_xlsx(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.xls'):
            file_path = os.path.join(directory, filename)
            try:
                print(f"Converting {filename}...")
                convert_file(file_path)
                print(f"Successfully converted {filename}.")
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")


def convert_file(file_path):
    xls_wb = xlrd.open_workbook(file_path)
    new_wb = openpyxl.Workbook()
    sheets = xls_wb.sheets()

    for index, sheet in enumerate(sheets):
        if index == 0:
            new_ws = new_wb.active
            new_ws.title = sheet.name
        else:
            new_ws = new_wb.create_sheet(sheet.name)

        for rowIndex in range(sheet.nrows):
            row_values = sheet.row_values(rowIndex)
            row_values = convert_date_cells(
                row_values, sheet, rowIndex, xls_wb.datemode)
            for colIndex, value in enumerate(row_values):
                new_ws.cell(rowIndex + 1, colIndex + 1).value = value

    save_converted_file(new_wb, file_path)


def convert_date_cells(row_values, sheet, rowIndex, datemode):
    for colIndex, _ in enumerate(row_values):
        if sheet.cell_type(rowIndex, colIndex) == xlrd.XL_CELL_DATE:
            date_value = sheet.cell_value(rowIndex, colIndex)
            try:
                date = xlrd.xldate.xldate_as_datetime(date_value, datemode)
                formatted_date = date.strftime("%d/%m/%Y")
                row_values[colIndex] = formatted_date
            except xlrd.xldate.XLDateError as e:
                print(
                    f"Error converting date at row {rowIndex+1}, column {colIndex+1}: {e}")
    return row_values


def save_converted_file(new_wb, original_file_path):
    new_file_path = original_file_path[:-4] + '.xlsx'
    new_wb.save(new_file_path)
    print(f"Saved converted file to {new_file_path}")


# Specify the directory containing your xls files
directory = 'C:/Wineflow'
convert_xls_to_xlsx(directory)
