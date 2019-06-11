import os
from pathlib import Path

cwd = Path(os.getcwd())
week_to_process = "02"

if not os.path.exists("Results"):
    os.makedirs("Results")

if not os.path.exists(f"Results/Week_{week_to_process}"):
    os.makedirs(f"Results/Week_{week_to_process}")

if os.path.isfile(f"Results/Week_{week_to_process}/01.xlsx"):
    os.remove(f"Results/Week_{week_to_process}/01.xlsx")

print(Path(f"{os.getcwd()}" + "\Results"))