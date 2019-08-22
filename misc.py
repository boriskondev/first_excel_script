# TODO: everything is working fine. Just emails need to be implemented.

import pandas as pd

codes_to_remove = ["АВТ.КОД/АС:", "Авт. код: ", "Авт. код ", "Авт.код :", "Авт. код. ", "АВТ. КОД/ ", "АВТ.КОД/AC:",
                   "AUTH.CODE:", "AUTH CODE:", "Auth. Code ", "PRE-AUTH ", "ACC", "AC:", "АС:", "AC :", "АС :", "AC: ",
                   "AC", "АС", "AС", "АC", "AC ", "ac ", "Ас ", "ac", "Ac", "tid "]

test_dataframe = pd.read_excel("template.xlsx")

# test_dataframe.columns = test_dataframe.columns.str.strip()
# test_dataframe["Име*:"] = test_dataframe["Име*:"].str.strip()


for code in codes_to_remove:
    test_dataframe["Отор. код на ПОС бележка*:"] = [x.strip().replace(code, '').strip() for x in test_dataframe["Отор. код на ПОС бележка*:"]]

test_dataframe.to_excel("template_clean.xlsx", index=False)

# For removing emails
# indices_to_remove = try_df[(try_df["Имейл:"] == "boris@gmail.com")].index
# try_df.drop(indices_to_remove, inplace=True)
