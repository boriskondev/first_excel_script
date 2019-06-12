from pathlib import Path
import pandas as pd
from datetime import datetime
import re

time_start = datetime.now()

pattern = r"\d{2}\.\d{2}\.\d{4}"

source_path = Path("C:/Users/pmg23_b.kondev/Desktop/Python/all_in_one_script/Test folder")

week_to_process = int(input("Week to process: "))

all_data_frames = []
week_dates_list = []

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
            print(f"{new_col_name}", end=" ")
            break

all_data = pd.concat(all_data_frames)
all_data = all_data.reset_index(drop=True)

values_to_skip = ["ID", "transaction_id", "TransactionID", "TRNX ID", "TRANSACTION ID", "Card Serno", "Transaction ID",
                  "Source Reg Num", "id", "ID of the transaction", "transaction_number", "Id", "TRNX No. ", "Id на транзакция"]

for value_to_skip in values_to_skip:
    indices_to_remove = all_data[(all_data["Transaction"] == value_to_skip)].index
    all_data.drop(indices_to_remove, inplace=True)

winners_to_draw = int(input("\nWinners to draw: "))
reserves_to_draw = int(input("Reserves to draw: "))
entries_to_sample = winners_to_draw + reserves_to_draw
sample_data_frame = all_data.sample(n=entries_to_sample)

winners = sample_data_frame.iloc[:winners_to_draw]
reserves = sample_data_frame.iloc[winners_to_draw:]
print(f"Winner/s: \n {winners}")
print(f"Reserve/s: \n {reserves}")

file_with_all_data = f"C:/Users/pmg23_b.kondev/Desktop/Python/all_in_one_script/Test folder/All_week_{week_to_process}.csv"
all_data.to_csv(file_with_all_data, index=False)

week_dates_list = all_data["Date"].unique()
for date in week_dates_list:
    date_data_frame = all_data[all_data.Date == date]
    date_file_to_save = f"C:/Users/pmg23_b.kondev/Desktop/Python/all_in_one_script/Test folder/{date}.csv"
    date_data_frame.to_csv(date_file_to_save, index=False, header=None)
    print(f"{date}", end=" ")

sample_data_frame.iloc[0] = ["Winners", None, None]
sample_data_frame.iloc[reserves_to_draw] = ["Reserves", None, None]
file_with_all_drawn = f"C:/Users/pmg23_b.kondev/Desktop/Python/all_in_one_script/Test folder/Winners_week_{week_to_process}.csv"
sample_data_frame.to_csv(file_with_all_drawn, index=False, header=None)

print("\nAll files are ready!")

time_elapsed = datetime.now()
time_took = time_elapsed - time_start
print(f"The execution of this script took {time_took.seconds} seconds.")
