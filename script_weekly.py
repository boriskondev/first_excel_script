import docx
from datetime import datetime
import pandas as pd
from pathlib import Path
import os

time_start = datetime.now()

project_config = docx.Document("config.docx")

weekly_source_path = Path(project_config.paragraphs[1].text.split("=")[1].replace("\"", ""))
output_path = Path(project_config.paragraphs[2].text.split("=")[1].replace("\"",""))

os.chdir(output_path)

week_to_process = int(input("Week to process: "))
winners_to_draw = int(input("Winners: "))
reserves_to_draw = int(input("Reserves: "))

entries_to_sample = winners_to_draw + reserves_to_draw

if week_to_process < 10:
    if not os.path.exists(f"Week_0{week_to_process}"):
        os.makedirs(f"Week_0{week_to_process}")
else:
    if not os.path.exists(f"Week_{week_to_process}"):
        os.makedirs(f"Week_{week_to_process}")

for week_folder in weekly_source_path.glob("*"):
    if int(week_folder.name.split(".")[0]) == week_to_process:
        for file in week_folder.iterdir():
            if file.suffix == ".xlsx":
                data_frame = pd.read_excel(file, sheet_name=0, header=None, skiprows=None)
                weekly_data_frame = data_frame.sample(n=entries_to_sample)
                whole_week = week_folder.name.split()[1]
                if week_to_process < 10:
                    weekly_winners_file = f"Week_0{week_to_process}/Winners_week_{whole_week}.xlsx"
                else:
                    weekly_winners_file = f"Week_{week_to_process}/Winners_week_{whole_week}.xlsx"
                if os.path.exists(weekly_winners_file):
                    os.remove(weekly_winners_file)
                weekly_data_frame.to_excel(weekly_winners_file, index=False, header=None)

time_end = datetime.now()
time_took = time_end - time_start

print(f"The execution of this script took {time_took.seconds} seconds.")