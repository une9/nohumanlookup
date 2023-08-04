

from util.fetchColumnInfo import fetchColumnInfo
from util.getTargetCols import getTargetCols, getTargetTableAndCols
from util.getQuery import getQuery
from util.getEndRow import getEndRow
from util.lookupInfoAndShowResult import lookupInfoAndShowResult
from util.getFile import getSrcFile
from pprint import pprint


# read Excel
target_wb, target_sheet = getSrcFile()

# get Query to parse
# end_row, query = getQuery(target_wb, target_sheet)

end_row = getEndRow(target_wb, target_sheet)

if end_row is None:   
    print("문서 양식을 확인해주세요")
else:
    # print(f"!!!!!! query: {query}")
    ### SQL parsing
    # target_cols = getTargetCols(query)
    target_cols = getTargetTableAndCols(target_wb, target_sheet, end_row)

    ### fetching column information from DB
    column_info_dict = fetchColumnInfo(target_cols)
    pprint(column_info_dict)

    ### write result of comparing on new excel sheet ('비교')
    lookupInfoAndShowResult(target_wb, target_sheet, end_row, column_info_dict)

