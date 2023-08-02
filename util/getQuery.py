import xlwings as xw


def getQuery(wb, sheet):
    # read excel and get Query
    target_text = 'sample'

    for row in sheet.range("A:A"):
        if row.value == target_text:
            # Get the cell on the right (column B)
            target_row_number = row.row
            return target_row_number, row.offset(0, 1).value

    return None, None
