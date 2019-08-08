from pathlib import Path
import pandas as pd
import docx
import yaml

rename_folder = Path("C:/Users/pmg23_b.kondev/Desktop/Python/lottery_script/Rename")

project_config = docx.Document("config.docx")

banks_list = yaml.safe_load(project_config.paragraphs[3].text.split("=")[1].replace("\"",""))

for file in rename_folder.glob("*"):
    current_data_frame = pd.read_excel(file, sheet_name=0, header=None, skiprows=None)
    current_data_frame.rename(columns={0: "Bank", 1: "Transaction", 2: "Date"}, inplace=True)
    current_data_frame["Bank"].replace(banks_list, inplace=True)
    current_data_frame["Bank"] = current_data_frame["Bank"].map(lambda x: x.encode("utf-8").decode("utf-8"))
    current_data_frame.to_excel("Try1.xlsx", index=False)
    # not working with .csv
