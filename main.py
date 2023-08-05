import sys

from util.fetchColumnInfo import fetchColumnInfo
from util.getTargetCols import getTargetTableAndCols
from util.lookupInfoAndShowResult import lookupInfoAndShowResult
from util.getFileNameList import getFileNameList
from util.getFileProperties import getFileProperties
from pprint import pprint


def main():
    # Check if at least one parameter was provided
    param = None
    if len(sys.argv) >= 2:
        param = sys.argv[1]

    # get target Files
    files = getFileNameList() if param is None else [param]

    for file_name in files:
        print(f"[START] {file_name} -------------")
        # (wb, ws, endRow)
        target_wb, target_ws, end_row = getFileProperties()

        if end_row is None:   
            print("문서 양식을 확인해주세요")
        else:
            target_cols = getTargetTableAndCols(target_wb, target_ws, end_row)

            ### fetching column information from DB
            column_info_dict = fetchColumnInfo(target_cols)
            pprint(column_info_dict)

            ### write result of comparing on new excel sheet ('비교')
            lookupInfoAndShowResult(target_wb, target_ws, end_row, column_info_dict)

        target_wb.close()
        print(f"[END] {file_name} -------------")





if __name__ == "__main__":
    main()