from pathlib import Path
import pandas as pd

rename_folder = Path("C:/Users/pmg23_b.kondev/Desktop/Python/first_python_script/Rename")

banks = {"all": "Алианц Банк България",
         "dsk": "Банка ДСК",
         "bac": "Българо-Американска Кредитна Банка",
         "bia": "Бяла Карта",
         "pir": "Пиреос Банк",
         "inv": "Инвестбанк АД",
         "ubb": "Обединена Българска Банка",
         "exp": "Експресбанк АД",
         "pro": "ПроКредит Банк",
         "fib": "Първа инвестиционна банка",
         "rai": "Райфайзенбанк България",
         "tex": "Тексим Банк",
         "ucb": "УниКредит Булбанк",
         "ccb": "Централна Кооперативна Банка",
         "pos": "Юробанк България АД"}

for file in rename_folder.glob("*"):
    current_data_frame = pd.read_excel(file, sheet_name=0, header=None, skiprows=None)
    current_data_frame.rename(columns={0: "Bank", 1: "Transaction", 2: "Date"}, inplace=True)
    current_data_frame["Bank"].replace(banks, inplace=True)
    current_data_frame["Bank"] = current_data_frame["Bank"].map(lambda x: x.encode("utf-8").decode("utf-8"))
    current_data_frame.to_excel("Try1.xlsx", index=False)
    # not working with .csv

