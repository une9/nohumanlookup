import xlwings as xw


def getEndRow(wb, sheet):
    # read excel and get End Row
    target_text = 'sample'

    for row in sheet.range("A:A"):
        if row.value == target_text:
            # Get the cell on the right (column B)
            target_row_number = row.row
            return target_row_number - 3

    return None
