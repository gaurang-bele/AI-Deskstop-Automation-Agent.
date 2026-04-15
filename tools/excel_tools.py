import pandas as pd

FILE = "data.xlsx"

def delete_rows(n):

    df = pd.read_excel(FILE)

    df = df.iloc[n:]

    df.to_excel(FILE, index=False)

    return f"{n} rows deleted successfully"