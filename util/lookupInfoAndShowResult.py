import xlwings as xw

def prepareSheetForComparison(wb, ws, comparison_sheet_name):
    comparison_sheet = None
    # Check if the sheet 'comparison' already exists in the workbook
    for sht in wb.sheets:
        if sht.name == comparison_sheet_name:
            comparison_sheet = sht
            break

    # If 'comparison' sheet doesn't exist, create a new one
    if comparison_sheet is None:
        comparison_sheet = wb.sheets.add(comparison_sheet_name, after=ws)
    
    return comparison_sheet


def copySourceData(source_sheet, comparison_sheet, end_row):
    # Specify the source range (C4:F41) and target range on the 'comparison' sheet
    print(f"copy source data: {source_sheet, comparison_sheet, end_row}")
    source_range = source_sheet.range(f'K15:Q{end_row}')
    target_range = comparison_sheet.range('B4')

    # Copy the data from the source range to the target range
    comparison_sheet.range('B2').value = "기존 문서"
    source_range.api.Copy()
    target_range.api.PasteSpecial()

    # Clear the clipboard to avoid Excel freezing issues
    xw.apps.active.api.CutCopyMode = 0

    # change cell color to default (None)
    comparison_sheet.range(f'B5:H{end_row}').color = None
    return


def createFetchDataTable(comparison_sheet, column_info_dict, end_row):
    # create table header by copy
    comparison_sheet.range('J2').value = "실제 DB"
    header_source_range = comparison_sheet.range(f'B4:H{end_row-11}')
    header_target_range = comparison_sheet.range('J4')
    header_source_range.api.Copy()
    header_target_range.api.PasteSpecial()

    # delete table values
    comparison_sheet.range(f'J5:P{end_row-11}').value = ""

    # Clear the clipboard to avoid Excel freezing issues
    xw.apps.active.api.CutCopyMode = 0
    return


def isSame(a, b):
    if type(a) is float:
        a = int(a)
    if type(b) is float:
        b = int(b)

    a, b = str(a).strip().upper(), str(b).strip().upper()
    strings = ("VARCHAR", "VARCHAR2")
    numbers = ("NUMBER", "INT", "DECIMAL")
    blanks = ("", None, "NONE")

    if a == b:
        return True
    elif a in strings and b in strings:
        return True
    elif a in numbers and b in numbers:
        return True
    elif a in blanks and b in blanks:
        return True

    # print(f"a: {a} / b: {b}")
    return False


def doLookup(sheet, column_info_dict, end_row):
    # src_eng_name_col, src_kor_name_col, src_col_type_col, src_col_size_col, src_not_null_col, src_pk_col = "B", "C", "E", "F", "G", "H"
    # db_eng_name_col, db_kor_name_col, db_col_type_col, db_col_size_col, db_not_null_col, db_pk_col = "J", "K", "M", "N", "O", "P"

    # print(f"******-----------")
    # print(f"******-----------{column_info_dict}")

    src_cols = ["B", "C", "E", "F", "G", "H"]
    db_cols = ["J", "K", "M", "N", "O", "P"]
    note_col = "Q"

    data_mapping = {
        "J" : "COLUMN_NAME",
        "K" : "COLUMN_COMMENT",
        "M" : "DATA_TYPE",
        "N" : "CHARACTER_MAXIMUM_LENGTH",
        "N2" : "NUMERIC_PRECISION",      # decimal type인 경우 CHARACTER_MAXIMUM_LENGTH 대신 사용
        "O" : "IS_NOT_NULLABLE",    # 필수여부
        "P" : "IS_PRIMARY_KEY",    # PK
    }
    # TABLE_NAME, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, COLUMN_COMMENT, IS_NOT_NULLABLE, IS_PRIMARY_KEY

    table_row_color = (128, 128, 128)   # grey
    alert_color = (255, 0, 0)        # red
    warning_color = (255, 255, 0)    # yellow  

    table = None
    for row in range(5, end_row + 1):
        # eng_name, kor_name, col_type, col_size, not_null, pk = 
        r = str(row)
        src_key_cell = sheet.range(src_cols[0] + r)
        db_key_cell = sheet.range(db_cols[0] + r)
        src_cells = sheet.range(f"{src_cols[0]}{r}:{src_cols[-1]}{r}")
        db_cells = sheet.range(f"{db_cols[0]}{r}:{db_cols[-1]}{r}")

        # table_names = list(map(lambda x: x.upper(), column_info_dict.keys()))
        table_names = list(column_info_dict.keys())
        # print(f"table names: {table_names}")
        
        src_key = src_key_cell.value

        if src_key is None:                        # 빈 값일 때
            continue
        elif src_key.upper() in table_names:       # 테이블명일 때
            src_cells.color = table_row_color
            # db_cells.color = table_row_color
            src_cells.api.Copy()
            db_key_cell.api.PasteSpecial()

            # Clear the clipboard to avoid Excel freezing issues
            xw.apps.active.api.CutCopyMode = 0

            # set table
            table = src_key
        else:                                                 # 컬럼명일 때
            # print(f"-----------table: {table}")
            # print(f"-----------{column_info_dict[table]}")
            # print(f"-----------{column_info_dict[table].keys()}")
            # print(f"-----------{src_key}")
            if src_key not in column_info_dict[table].keys():
                db_key_cell.value = src_key
                sheet.range(db_cols[1] + r).value = "(정보 없음)"
                db_cells.color = warning_color
                print(f"!!! 정보 없음 - {src_key}")
                continue

            target_info = column_info_dict[table][src_key]
            # print(f"target_info : {target_info}")
            if len(target_info) == 0:                   # 일치하는 정보가 없을 때
                print(f"***No target Info : table - {table} / col - {src_key}")
            else:                                       # 일치하는 정보가 있을 때 -> 기존 문서 데이터와 비교
                for i in range(len(db_cols)):
                    target_info_key = data_mapping[db_cols[i]]
                    val = target_info[target_info_key]
                    if val is None and i > 0 and target_info[data_mapping[db_cols[i-1]]]:    # decimal 타입인 경우
                        val = target_info[data_mapping["N2"]]
                    target_cell = sheet.range(db_cols[i] + r)
                    src, tgt = sheet.range(src_cols[i] + r).value, val
                    target_cell.value = tgt
                    if src is None and tgt != '':
                        target_cell.color = warning_color
                    elif not isSame(src, tgt):
                        target_cell.color = alert_color
                        
    return

    


def lookupInfoAndShowResult(wb, ws, query_row, column_info_dict):
    comparison_sheet_name = "비교"
    comparison_sheet = prepareSheetForComparison(wb, ws, comparison_sheet_name)
    print(f"comparision sheet : {comparison_sheet}")

    end_row = query_row - 3
    copySourceData(ws, comparison_sheet, end_row)

    createFetchDataTable(comparison_sheet, column_info_dict, end_row)

    doLookup(comparison_sheet, column_info_dict, end_row)

    return