# Remove duplicates functionality to add some time

import pandas as pd

database = pd.read_excel("script.xlsx", sheet_name=0, skiprows=None)
duplicates = database[database.duplicated(["Отор. код на ПОС бележка*: ", "Имейл*: ", "Дата на ПОС плащане*:"], keep="first")]
duplicates.to_excel("duplicates.xlsx", index=False)
database.drop_duplicates(["Отор. код на ПОС бележка*: ", "Имейл*: ", "Дата на ПОС плащане*:"], keep="first", inplace=True)
database.to_excel("script_no_duplicates.xlsx", index=False)