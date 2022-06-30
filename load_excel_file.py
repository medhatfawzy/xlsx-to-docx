import pandas as pd
from coordinate import coordination

CODE_COLUMNS = ["المخزن", "بلوك", "راك", "رف"]

def load_excel_file(path):
    data = pd.read_excel(path)
    data["coordinate"] = data[CODE_COLUMNS].apply(coordination, axis=1)
    data.drop(CODE_COLUMNS, inplace=True, axis=1)
    return data.values

