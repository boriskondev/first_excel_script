from pathlib import Path
from os import chdir
import pandas as pd

banks = {
    "DSK": "dsk",
    "Postbank": "pos",
    "Allianz": "all",
}

source_path = Path("C:/Users/pmg23_b.kondev/Desktop/Python/first_python_script/Test folder/bac")

data_frames = []
all_data_frames = []

for folder in source_path.glob("*"):
    for file in folder.iterdir():
        data_frame = pd.read_excel(file, sheet_name=0, usecols="A", header=None, skiprows=1)
        new_col_name = folder.name
        data_frame.insert(0, "bank", new_col_name)
        data_frames.append(data_frame)

all_data = pd.concat(data_frames)
chdir(source_path)
path = Path() / '_Results'
path.mkdir()
all_data.to_excel("_Results/Results.xlsx", index=False, sheet_name="All_data")