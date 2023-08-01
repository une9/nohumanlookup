import xlwings as xw

def prepareSheetForComparison(wb):
    comparison_sheet = None
    # Check if the sheet 'comparison' already exists in the workbook
    for sht in wb.sheets:
        if sht.name == 'comparison':
            comparison_sheet = sht
            break

    # If 'comparison' sheet doesn't exist, create a new one
    if comparison_sheet is None:
        comparison_sheet = wb.sheets.add('comparison')
    
    return comparison_sheet


def copySourceData(source_sheet, comparison_sheet, end_row):
    # Specify the source range (C4:F41) and target range on the 'comparison' sheet
    source_range = source_sheet.range(f'C4:G{end_row - 1}')
    target_range = comparison_sheet.range('C4')

    # Copy the data from the source range to the target range
    source_range.api.Copy()
    target_range.api.PasteSpecial()

    # Clear the clipboard to avoid Excel freezing issues
    xw.apps.active.api.CutCopyMode = 0


def writeFetchData(column_info_dict):
    return


def doLookup(column_info_dict):
    return


def lookupInfoAndShowResult(wb, sheet, end_row, column_info_dict):
    comparison_sheet = prepareSheetForComparison(wb)

    copySourceData(sheet, comparison_sheet, end_row)

    writeFetchData(column_info_dict)

    doLookup()

    return