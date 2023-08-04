from sql_metadata import Parser
from collections import defaultdict
import xlwings as xw
from pprint import pprint


def getTargetCols(query):
    columns = Parser(query).columns
    # print("** tables: ", Parser(query).tables)
    pprint(columns)

    result = defaultdict(set)

    try:
        for column in columns:
            table_name, col_name = column.split(".")
            result[table_name].add(col_name)
    except:
        print(f"Exception: {column}")
        

    return result


def getTargetTableAndCols(wb, ws, end_row):
    resultDict = defaultdict(set)
    table = None

    for i in range(16, end_row + 1):
        val = ws.range(f"K{i}").value
        if val is None:                         # 빈칸일 때
            continue

        if ws.range(f"N{i}").value == "TABLE":  # 테이블명일 때
            table = val
        else:                                   # 컬럼명일 때
            resultDict[table].add(val)
    
    pprint(resultDict)

    return resultDict