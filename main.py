

from util.fetchColumnInfo import fetchColumnInfo
from util.getTargetCols import getTargetCols
from util.getQuery import getQuery
from util.lookupInfoAndShowResult import lookupInfoAndShowResult
from util.getFile import getSrcFile



# read Excel
target_wb, target_sheet = getSrcFile()

# get Query to parse
end_row, query = getQuery(target_wb, target_sheet)

if query is None:   
    print("target query is not found in the sheet.")
else:
    # print(f"!!!!!! query: {query}")
    ### SQL parsing
    target_cols = getTargetCols(query)

    ### fetching column information from DB
    column_info_dict = fetchColumnInfo(target_cols)
    print(column_info_dict)

    ### write result of comparing on new excel sheet ('비교')
    lookupInfoAndShowResult(target_wb, target_sheet, end_row, column_info_dict)

