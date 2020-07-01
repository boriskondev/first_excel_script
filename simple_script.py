# A really simple one, but with real time information what is going on during execution.

import pandas as pd
import os

small_winners = 10000
big_winners = 10

input_file = ""  # Enter file name here
output_all_submissions = "all_submissions_week2.xlsx"
output_winners_xlsx = "all_winners_week2.xlsx"
output_winners_txt = "winners_week2.txt"


first_sheet = pd.read_excel(input_file, sheet_name=0, skiprows=None)
print(f"working on 1st sheet with {first_sheet.shape[0]} entries inside...")

second_sheet = pd.read_excel(input_file, sheet_name=1, skiprows=None)
print(f"working on 2nd sheet with {second_sheet.shape[0]} entries inside...")

print("generating result...")
total_submissions = pd.concat([first_sheet, second_sheet])
total_submissions.columns = ["Winners"]

print("drawing winners...")
winners = total_submissions.sample(n=small_winners + big_winners)

print("saving the result...")
total_submissions.to_excel(output_all_submissions, index=False, header=False)
winners.to_excel(output_winners_xlsx, index=False, header=False)

print("preparing the notary file...")
winners_list = winners["Winners"].values.tolist()
winners_list = list(map(lambda x: str(x), winners_list))
winners_list_str = ", ".join(winners_list)

print(f"saving notary file with {len(winners_list)} winners in it...")
with open(output_winners_txt, "w") as file:
    file.write(winners_list_str)

print("done!")