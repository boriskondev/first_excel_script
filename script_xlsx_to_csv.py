from pathlib import Path
import pandas as pd
from datetime import datetime
import re

time_start = datetime.now()

pattern = r"\d{2}\.\d{2}\.\d{4}"

source_path = Path("C:/Users/pmg23_b.kondev/Desktop/Python/first_python_script/Test folder")

week_to_process = int(input("Week to process: "))

all_data_frames = []

for bank_folder in source_path.glob("*"):
    new_col_name = bank_folder.name
    for week_folder in bank_folder.iterdir():
        if int(week_folder.name.split(".")[0]) == week_to_process:
            for file in week_folder.iterdir():
                current_data_frame = pd.read_excel(file, sheet_name=0, usecols="A", header=None, skiprows=None)
                current_data_frame.rename(columns={0: "Transaction"}, inplace=True)
                current_data_frame.insert(0, "Bank", new_col_name)
                searched = re.search(pattern, file.name)
                current_data_frame.insert(2, "Date", searched[0])
                all_data_frames.append(current_data_frame)
            print(f"Ready {new_col_name}")
            break

values_to_skip = ["ID", "transaction_id", "TransactionID", "TRNX ID", "TRANSACTION ID", "Card Serno",
                  "Transaction ID", "Source Reg Num", "id", "ID of the transaction", "transaction_number", "Id"]

all_data = pd.concat(all_data_frames)

for value_to_skip in values_to_skip:
    indices_to_remove = all_data[(all_data["Transaction"] == value_to_skip)].index
    all_data.drop(indices_to_remove, inplace=True)

all_data.to_csv(f"C:/Users/pmg23_b.kondev/Desktop/Python/first_python_script/Test folder/Results_week_{week_to_process}.csv", index=False)

time_elapsed = datetime.now()
time_took = time_elapsed - time_start
print(f"The execution of this script {time_took.seconds} seconds.")
