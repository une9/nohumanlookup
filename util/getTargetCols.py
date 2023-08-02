from sql_metadata import Parser
from collections import defaultdict


def getTargetCols(query):
    columns = Parser(query).columns
    print("** tables: ", Parser(query).tables)
    print(columns)

    result = defaultdict(set)

    try:
        for column in columns:
            table_name, col_name = column.split(".")
            result[table_name].add(col_name)
    except:
        print(f"Exception: {column}")
        

    return result