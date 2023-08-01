import xlwings as xw

from util.fetchColumnInfo import fetchColumnInfo
from util.getTargetCols import getTargetCols
from util.getQuery import getQuery
from util.lookupInfoAndShowResult import lookupInfoAndShowResult

# read Excel
file_name = "./sample_document.xlsx"
target_wb = xw.Book(file_name, read_only=True)
target_sheet = target_wb.sheets['Sheet1']

end_row, query = getQuery(target_wb, target_sheet)

if query is None:
    print("target query is not found in the sheet.")
else:
    # print(f"!!!!!! query: {query}")
    ### SQL parsing
    target_cols = getTargetCols(query)

    ### fetching column information
    column_info_dict = fetchColumnInfo(target_cols)
    print(column_info_dict)

    ### write result of comparing on new excel sheet ('comparison')
    # lookupInfoAndShowResult(target_wb, target_sheet, end_row, column_info_dict)

