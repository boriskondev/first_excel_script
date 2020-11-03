from datetime import datetime
from pathlib import Path
import docx
import pandas as pd
import os
import numpy as np
import yaml

# time_start = datetime.now()

project_config = docx.Document("config_new.docx")

source_path = Path(project_config.paragraphs[0].text.split("=")[1].replace("\"", ""))  # pathlib.WindowsPath
output_path = Path(project_config.paragraphs[1].text.split("=")[1].replace("\"", ""))   # pathlib.WindowsPath
banks_names = yaml.safe_load(project_config.paragraphs[2].text.split("=")[1].replace("\"", ""))  # dict
codes_to_remove = project_config.paragraphs[3].text.split("=")[1].strip().replace("\"", "").split(", ")  # list

statistics = []

week_to_process = 1
daily_winners = 5
daily_reserves = 5
weekly_winners = 1
weekly_reserves = 5

whole_week = ""

os.chdir(output_path)

for element in source_path.glob("*"):
    if os.path.isdir(element):
        if element.name != "_Results" and int(element.name.split(".")[0]) == week_to_process:
            for file in element.iterdir():
                if file.suffix == ".xlsx":

                    all_df = pd.read_excel(file, sheet_name=0, skiprows=None)
                    all_df.columns = all_df.columns.str.strip()
                    all_df["Банка*:"] = all_df["Банка*:"].str.strip()
                    all_df["Reg_date"], all_df["Reg_time"] = all_df["Submitted date"].str.split(" ").str

                    all_df.to_excel("test.xlsx", index=False)
