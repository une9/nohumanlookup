from sql_metadata import Parser
from collections import defaultdict


def getTargetCols(query):
    columns = Parser(query).columns
    print(columns)

    result = defaultdict(set)

    for column in columns:
        table_name, col_name = column.split(".")
        result[table_name].add(col_name)

    return result