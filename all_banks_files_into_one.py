from pathlib import Path
import pandas as pd
from os import chdir
from datetime import datetime

time_start = datetime.now()
source_path = Path("C:/Users/pmg23_b.kondev/Desktop/Python/first_python_script/Test folder")

week_to_process = int(input("Week to process: "))

all_data_frames = []

for bank_folder in source_path.glob("*"):
    new_col_name = bank_folder.name
    for week_folder in bank_folder.iterdir():
        if int(week_folder.name.split(".")[0]) == week_to_process:
            for file in week_folder.iterdir():
                current_data_frame = pd.read_excel(file, sheet_name=0, usecols="A", header=None, skiprows=1)
                current_data_frame.rename(columns={0: "Transaction"}, inplace=True)
                current_data_frame.insert(0, "Bank", new_col_name)
                all_data_frames.append(current_data_frame)
            print(f"Ready {new_col_name}")
            break

all_data = pd.concat(all_data_frames)

chdir(source_path)
path = Path() / '_Results'
path.mkdir()
all_data.to_excel("_Results/Results.xlsx", index=False, sheet_name="All_data")

time_elapsed = datetime.now()
time_took = time_elapsed - time_start
print(f"The execution of this script {time_took.seconds} seconds.")

# The script works well with (almost) all the files of the banks.
# The ones of UniCredit should be additionally saved into .xlsx as there are fucked up somehow.
# Important:
# Dani - how do you handle UniCredit's files?
# Some values start on the first row, so it is now an option to ignore it - some other check here
