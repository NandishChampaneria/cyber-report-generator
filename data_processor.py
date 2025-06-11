import pandas as pd

def read_excel_data(path: str) -> dict:
    xls = pd.ExcelFile(path)
    data = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet)
        data[sheet] = df

    return data