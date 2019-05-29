from pathlib import Path

data_folder = Path("C:/Users/pmg23_b.kondev/Desktop/Files")

file_to_open = data_folder / "file.txt"

f = open(file_to_open, "r")
print(f.readlines(1))