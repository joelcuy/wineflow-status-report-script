import os
import xlrd
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

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
            row_values = convert_date_cells(row_values, sheet, rowIndex, xls_wb.datemode)
            for colIndex, value in enumerate(row_values):
                new_ws.cell(rowIndex + 1, colIndex + 1).value = value

    save_converted_file(new_wb, file_path)
    style_workbook(file_path[:-4] + '.xlsx')

def convert_date_cells(row_values, sheet, rowIndex, datemode):
    for colIndex, _ in enumerate(row_values):
        if sheet.cell_type(rowIndex, colIndex) == xlrd.XL_CELL_DATE:
            date_value = sheet.cell_value(rowIndex, colIndex)
            try:
                date = xlrd.xldate.xldate_as_datetime(date_value, datemode)
                formatted_date = date.strftime("%d/%m/%Y")
                row_values[colIndex] = formatted_date
            except xlrd.xldate.XLDateError as e:
                print(f"Error converting date at row {rowIndex+1}, column {colIndex+1}: {e}")
    return row_values

def save_converted_file(new_wb, original_file_path):
    new_file_path = original_file_path[:-4] + '.xlsx'
    new_wb.save(new_file_path)
    print(f"Saved converted file to {new_file_path}")

def style_workbook(file_path):
    wb = openpyxl.load_workbook(file_path)
    for ws in wb.worksheets:
        max_row = ws.max_row
        max_col = ws.max_column

        # Set header style
        header_font = Font(color="FFFFFF", bold=True)
        header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        header_border = Border(left=Side(style='thin', color='FFFFFF'),
                               right=Side(style='thin', color='FFFFFF'),
                               top=Side(style='thin', color='FFFFFF'),
                               bottom=Side(style='thin', color='FFFFFF'))
        alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = header_border
            cell.alignment = alignment

        ws.row_dimensions[1].height = 40  # Height for header row

        # Auto-fit columns and rows (approximation)
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 10)
            ws.column_dimensions[column].width = adjusted_width



        # Apply striped pattern to content rows
        odd_fill = PatternFill(start_color="B5E6A2", end_color="B5E6A2", fill_type="solid")
        even_fill = PatternFill(start_color="DAF2D0", end_color="DAF2D0", fill_type="solid")

        for row in range(2, max_row + 1):
            fill = even_fill if row % 2 == 0 else odd_fill
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.fill = fill
                cell.border = header_border

        # Freeze the top row
        ws.freeze_panes = ws['A2']

        # Add table-like appearance
        ws.auto_filter.ref = ws.dimensions

    wb.save(file_path)
    print(f"Styled and saved workbook {file_path}")

# Specify the directory containing your xls files
directory = 'C:/Wineflow'
convert_xls_to_xlsx(directory)
