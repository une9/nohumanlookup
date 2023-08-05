from jproperties import Properties
import xlwings as xw


def getWorkbookAndWorkSheet(file_name):
    try:
        src_wb = xw.Book(file_name, read_only=True)
        src_ws = src_wb.sheets[2]

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


