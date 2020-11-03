from datetime import datetime
from pathlib import Path
import docx
import pandas as pd
import os
import numpy as np
import yaml

# time_start = datetime.now()

project_config = docx.Document("config_new.docx")

source_path = Path(project_config.paragraphs[0].text.split("=")[1].replace("\"", ""))  # <class 'pathlib.WindowsPath'>
output_path = Path(project_config.paragraphs[1].text.split("=")[1].replace("\"", ""))   # <class 'pathlib.WindowsPath'>
banks_names = yaml.safe_load(project_config.paragraphs[2].text.split("=")[1].replace("\"", ""))  # <class 'dict'>
