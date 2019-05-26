import pandas as pd
import re
from pathlib import Path
import os

banks = {"Allianz": "all", "Test": "tst"}

os.chdir("C:/Users/pmg23_b.kondev/Desktop/Files/Test")
data_folder = Path("C:/Users/pmg23_b.kondev/Desktop/Files/Test")

data_frames = []

for file in os.listdir(data_folder):
    opened_file = data_folder / file
    data_frame = pd.read_excel(opened_file, sheet_name=0, usecols="A", header=None, skiprows=1)
    folder_name = Path(data_folder).resolve().stem
    if folder_name in banks:
        new_col_name = banks[folder_name]
        data_frame.insert(0, "bank", new_col_name)
        data_frames.append(data_frame)

all_sheets = pd.concat(data_frames)
all_sheets.to_excel("results.xlsx", index=False, sheet_name="All_data")

# a working script for adding all files from a folder into one file
