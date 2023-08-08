from collections import defaultdict
import xlwings as xw
from pprint import pprint


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