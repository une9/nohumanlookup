from jproperties import Properties
import pymysql
from collections import defaultdict


def fetchColumnInfo(target_cols):
    # make db connection
    config = Properties()
    with open('db.properties', 'rb') as dbConfig:
        config.load(dbConfig, "utf-8")

    target_db = pymysql.connect(
        user=config.get('user').data, 
        passwd=config.get('passwd').data, 
        host=config.get('host').data, 
        db=config.get('db').data, 
        charset=config.get('charset').data
    )

    cursor = target_db.cursor(pymysql.cursors.DictCursor)

    fetching_query = """
                    SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, COLUMN_COMMENT, NUMERIC_PRECISION,
                            IF(IS_NULLABLE = 'NO', 'Y', '') AS IS_NOT_NULLABLE, 
                            IF(COLUMN_KEY = 'PRI', 'Y', '') AS IS_PRIMARY_KEY
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE 
                        TABLE_NAME = %s
                        AND COLUMN_NAME = %s;
                """
    
    column_info_dict = defaultdict(dict)
    
    for table, columns in target_cols.items():
        for column in columns:
            # print(table, " ", column)
            cursor.execute(fetching_query, (table, column))
            fetch_result = cursor.fetchall()
            if len(fetch_result) > 0:
                data = fetch_result[0]
                # print(table, " ", column, " ", fetch_result)
                column_info_dict[table][data['COLUMN_NAME']] = data

    return column_info_dict