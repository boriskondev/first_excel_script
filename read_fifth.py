import pandas as pd
import numpy as np

all_fifths = pd.DataFrame()

all_emails = pd.read_excel("test.xlsx", sheet_name=0, skiprows=None)

unique_emails = pd.read_excel("test.xlsx", sheet_name=0, skiprows=None)

unique_emails.drop_duplicates(["Имейл*:"], keep="first", inplace=True)

for email in unique_emails["Имейл*:"]:
    all_selected = all_emails[all_emails["Имейл*:"] == email]
    all_selected.index = np.arange(1, len(all_selected) + 1)
    each_fifth = all_selected[(all_selected.index % 5 == 0)]
    all_fifths = all_fifths.append(each_fifth)

print(all_fifths)
