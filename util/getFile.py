from jproperties import Properties
import xlwings as xw

def getSrcFile():
    config = Properties()
    with open('srcFile.properties', 'rb') as fileConfig:
        config.load(fileConfig, "utf-8")

    src_file = config.get('src_file').data
    src_sheet_name = config.get('src_sheet_name').data

    print(f"src_file: {src_file} / src_sheet_name: {src_sheet_name}")

    target_wb = xw.Book(src_file, read_only=True)
    target_sheet = target_wb.sheets[src_sheet_name]

    return target_wb, target_sheet