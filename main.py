

from util.fetchColumnInfo import fetchColumnInfo
from util.getTargetCols import getTargetCols
from util.getQuery import getQuery
from util.lookupInfoAndShowResult import lookupInfoAndShowResult
from util.getFile import getSrcFile



# read Excel
# file_name = "./sample_document.xlsx"
# src_file = "./SSGDX_D_아카데미_IF_PS_0025_과학관 전시이력및취소탭 바코드 센싱 자동조회_V0.6.xlsx"
# src_sheet_name = "IF_PS_0025"
# file_name = "./SSGDX_D_페이퍼리스_IF_MT_0588_인터페이스명세서 (점포마스터 이관) V0.6.xlsx"
# sheet_name = "IF_MT_0588"

# target_wb = xw.Book(src_file, read_only=True)
# target_sheet = target_wb.sheets[src_sheet_name]

target_wb, target_sheet = getSrcFile()
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
    lookupInfoAndShowResult(target_wb, target_sheet, end_row, column_info_dict)

