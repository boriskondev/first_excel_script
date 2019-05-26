import re

text = "27.08.2987"
pattern = r"\d{2}\.\d{2}\.\d{4}"

searched = re.findall(pattern, text)
print(sorted(searched))
