import os

folder_path = '/excels'

def getFileNameList():
    try:
        files = os.listdir(folder_path)
        excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]
        return excel_files
    except Exception as e:
        print(f"Error occurred: 파일 이름을 읽을 수 없습니다 - {e}")
        return None, None
    