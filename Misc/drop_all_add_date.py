import pandas as pd
from pathlib import Path
import re

values_to_skip = ["ID", "transaction_id", "TransactionID", "TRNX ID", "TRANSACTION ID", "Card Serno",
                  "Transaction ID", "Source Reg Num", "id", "ID of the transaction", "transaction_number", "Id"]

pattern = r"\d{2}\.\d{2}\.\d{4}"

data_folder = Path("C:/Users/pmg23_b.kondev/Desktop/Python/lottery_script/Test folder/all/01. 01.04-07.04.2019")

data_frames = []

for file in data_folder.glob("*"):
    data_frame = pd.read_excel(file, sheet_name=0, usecols="A", header=None, skiprows=None)
    data_frame.rename(columns={0: "Transaction"}, inplace=True)
    for value_to_skip in values_to_skip:
        indices_to_remove = data_frame[(data_frame["Transaction"] == value_to_skip)].index
        data_frame.drop(indices_to_remove, inplace=True)
    data_frame.insert(0, "Bank", "all")
    folder_name = file.name
    searched = re.search(pattern, folder_name)
    data_frame.insert(2, "Date", searched[0])
    data_frames.append(data_frame)

all_sheets = pd.concat(data_frames)
print(all_sheets)
print(all_sheets.shape)