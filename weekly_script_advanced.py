import docx
from datetime import datetime
import pandas as pd
from pathlib import Path
import os
import yaml

# Read the weekly file

test_df = pd.read_excel("test.xlsx", sheet_name=0, skiprows=None)

# Split the long date column into two separate columns - date and time of registration
test_df["Reg_date"], test_df["Reg_time"] = test_df["Submitted date"].str.split(" ").str

# Create a list with the unique values of the registration dates column
days_of_week = test_df["Reg_date"].unique()

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("test_multiple.xlsx", engine="xlsxwriter")

weekly_stats_file = "weekly_stats.txt"

# Filter the entries by day and save all in one file with different sheets
for day in days_of_week:
    day_df = test_df[test_df["Reg_date"] == day]
    with open(weekly_stats_file, "a+") as file:
        file.write(f"{day}: {day_df.shape[0]}\n")

    #day_df.to_excel(writer, sheet_name=day, index=False)

#writer.save()

#test_df.to_excel("test_ready.xlsx", index=False)

# import pandas as pd
# import numpy as np
#
# all_fifths = pd.DataFrame()
#
# all_emails = pd.read_excel("test4.xlsx", sheet_name=0, skiprows=None)
#
# unique_emails = pd.read_excel("test4.xlsx", sheet_name=0, skiprows=None)
#
# unique_emails.drop_duplicates(["Emails"], keep="first", inplace=True)
#
# for email in unique_emails["Emails"]:
#     all_selected = all_emails[all_emails["Emails"] == email]
#     all_selected.index = np.arange(1, len(all_selected) + 1)
#     each_fifth = all_selected[(all_selected.index % 5 == 0)]
#     all_fifths = all_fifths.append(each_fifth)
#
# #all_fifths = all_fifths.sample(n=10)
#
# #print("Winners:")
# print(all_fifths)
