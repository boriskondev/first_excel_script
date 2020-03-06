import docx
import pandas as pd
from pathlib import Path
import os
import numpy as np

project_config = docx.Document("config.docx")

weekly_source_path = Path(project_config.paragraphs[1].text.split("=")[1].replace("\"", ""))
output_path = Path(project_config.paragraphs[2].text.split("=")[1].replace("\"", ""))

codes_to_remove = ['AC ', 'ac', 'АС :', 'AUTH.CODE:', 'Ac', 'АВТ. КОД', 'Авт. Код : ', 'АВТ.КОД/АС:', 'ACC',
                   'AUTH CODE:', 'АВТОМ.КОД: ', 'АВТ. КОД: ', 'Авт.код', 'авт.код.', 'autn. Code', 'AC: ', 'ac ',
                   'tid ', 'AUTH.CODE : ', 'Авт. Код:', 'Авт. Код: ', 'AUTH. CODE:', 'Auth.code:', 'AUTH. CODE',
                   'Авт. код: ', 'Auth.code ', 'AC:', 'АC', 'АС:', 'АВТ КОД ', 'AUTH. CODE: ', 'Авт.код:', 'АВТ.КОД',
                   'АВТ.КОД. ', 'Авт. код. ', 'АВТ КОД', 'АВТ.КОД: ', 'AC', 'Auth. Code ', 'AС', 'авт. Код', 'АВТ. КОД:',
                   'Авт. код ', 'GPA.', 'АВТ.КОД:', 'авт. код ', 'AC :', 'Авт.код ', 'Ас ', 'АВТ. КОД/ ', 'PRE-AUTH ',
                   'Авт. Код ', 'авт.код:', 'АВТ.КОД/AC:', 'Авт.код :', 'АС', 'Авт. код/ ']

os.chdir(output_path)

week_to_process = 2

for weekly_folder_element in weekly_source_path.glob("*"):
    if os.path.isdir(weekly_folder_element) and weekly_folder_element.name != "_Results" \
            and int(weekly_folder_element.name.split(".")[0]) == week_to_process:
        for file in weekly_folder_element.iterdir():
            if file.suffix == ".xlsx":
                all_data_frame = pd.read_excel(file, sheet_name=0, skiprows=None)
                all_data_frame.columns = all_data_frame.columns.str.strip()
                all_data_frame["Firstname"] = all_data_frame["Firstname"].str.strip()

                for code in codes_to_remove:
                    all_data_frame["Transaction Code"] = \
                        [x.strip().replace(code, '').strip() for x in all_data_frame["Transaction Code"]]

                print(all_data_frame.shape)

                duplicates_data_frame = all_data_frame[
                    all_data_frame.duplicated(["Firstname", "Transaction Code"],
                                              keep="first")]
                all_data_frame.drop_duplicates(["Firstname", "Transaction Code"],
                                               keep="first", inplace=True)

                whole_week = weekly_folder_element.name.split()[1]

                weekly_duplicates = f"Week_{whole_week}_only_duplicates.xlsx"
                weekly_no_duplicates = f"Week_{whole_week}_without_duplicates.xlsx"
                weekly_fifths = f"Week_{whole_week}_fifths.xlsx"

                files = [weekly_duplicates, weekly_no_duplicates, weekly_fifths]

                for f in files:
                    if os.path.exists(f):
                        os.remove(f)

                duplicates_data_frame.to_excel(weekly_duplicates, index=False)
                all_data_frame.to_excel(weekly_no_duplicates, index=False)

                all_fifths = pd.DataFrame()

                all_data_frame["Email address (Personal)"] = all_data_frame["Email address (Personal)"].str.strip()
                emails_list = all_data_frame["Email address (Personal)"].unique()
                print(len(emails_list))

                for email in emails_list:
                    all_selected = all_data_frame[all_data_frame["Email address (Personal)"] == email]
                    all_selected.index = np.arange(1, len(all_selected) + 1)
                    each_fifth = all_selected[(all_selected.index % 5 == 0)]
                    all_fifths = all_fifths.append(each_fifth)

                all_fifths.to_excel(weekly_fifths, index=False)

                print(all_data_frame.shape)
                print(all_fifths.shape)

print("Done!")

# TO SPLIT THE DATES IN THE SHEETS!