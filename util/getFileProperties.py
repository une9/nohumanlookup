from jproperties import Properties
import xlwings as xw
import re

pattern = r'^IF_[A-Z]{2}_[0-9]{4}$'

def getWorkbookAndWorkSheet(file_name):
    try:
        src_wb = xw.Book(file_name, read_only=True)
        target_sheet_num = 2
        src_ws = src_wb.sheets[target_sheet_num]

        while not re.match(pattern, src_ws.name):
            target_sheet_num += 1
            src_ws = src_wb.sheets[target_sheet_num]

        print(f"src_file_name: {file_name} / src_sheet_name: {src_ws.name}")

        return src_wb, src_ws
    except Exception as e:
        print(f"Error occurred: {e}")
        return None, None
    

def getEndRow(ws):
    # read excel and get End Row
    target_text = 'sample'

    for row in ws.range("A:A"):
        if row.value == target_text:
            # Get the cell on the right (column B)
            target_row_number = row.row
            return target_row_number - 3

    return None


def getFileProperties(file_name):
    wb, ws = getWorkbookAndWorkSheet(file_name)
    endRow = getEndRow(ws)
    return wb, ws, endRow


