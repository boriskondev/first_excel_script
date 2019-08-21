# TODO: everything is working fine. Just emails need to be implemented.

'''
duplicates_data_drame = all_data_frame[
                    all_data_frame.duplicated(["Имейл*:", "Отор. код на ПОС бележка*:", "Дата на ПОС плащане*:"],
                                              keep="first")]
                all_data_frame.drop_duplicates(["Имейл*:", "Отор. код на ПОС бележка*:", "Дата на ПОС плащане*:"],
                                               keep="first", inplace=True)
'''

import pandas as pd

parts_to_remove = ["АВТ.КОД/АС:", "Авт. код: ", "ACC", "AC:", "АС:", "AC :", "АС :", "AC: ", "AC", "АС", "AС", "АC", "AC ", "ac ", "Ac"]

try_df = pd.read_excel("template.xlsx")

for to_remove in parts_to_remove:
    try_df["Отор. код на ПОС бележка*:"] = [x.strip().replace(to_remove, '').strip() for x in try_df["Отор. код на ПОС бележка*:"]]

try_df.to_excel("template_clean.xlsx", index=False)
