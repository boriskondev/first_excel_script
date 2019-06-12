import pandas as pd
from pathlib import Path

week_to_process = int(input("Week to process: "))
winners_to_draw = int(input("Winners per day: "))
reserves_to_draw = int(input("Reserves per day: "))

entries_to_sample = winners_to_draw + reserves_to_draw

source_path = Path("C:/Users/Boris Kondev/Desktop/Python/first_python_script/Weekly")

for week_folder in source_path.glob("*"):
    num = int(week_folder.name.split(".")[0])
    if int(week_folder.name.split(".")[0]) == week_to_process:
        for file in week_folder.iterdir():
            if file.suffix == ".xlsx":
                data_frame = pd.read_excel(file, sheet_name=0, header=None, skiprows=None)
                sample_data_frame = data_frame.sample(n=entries_to_sample)
                sample_data_frame.to_excel("Res.xlsx", index=False, header=None)

if week_to_process < 10:
    week_to_process = "0" + str(week_to_process)