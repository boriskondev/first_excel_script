import pandas as pd
import numpy as np

banks = ["ccb", "dsk", "exp", "ccb", "dsk", "dsk", "exp", "ccb", "ccb", "dsk"]
trans_ids = [389, 192, "d99", "IT", "DU&", 11, 130, 111, 11, 333]
dates = ["12.03", "13.03", "14.03", "12.03", "12.03", "12.03", "14.03", "13.03", "14.03", "14.03"]

data = {"Bank": banks, "Transaction": trans_ids, "Date": dates}

df = pd.DataFrame(data)

print(pd.pivot_table(df, index=["Bank"], values=["Transaction"], aggfunc=len, columns=["Date"]))