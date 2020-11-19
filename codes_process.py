import docx

new_codes = ['Авт, код', 'avt. Kod', 'AВТ КОД', 'ABT. КОД:']

# code = input().strip()
#
# while code != "END":
#     new_codes.append(code)
#     code = input().strip()
#
# print(new_codes)

project_config = docx.Document("config_new.docx")
codes_to_remove = project_config.paragraphs[3].text.split("=")[1].strip().replace("\"", "").split(", ")

all_codes = list(set(new_codes + codes_to_remove))
all_codes.sort()
all_codes.sort(key=len, reverse=True)

print(all_codes)

#  [print(f"{x} --> {len(x)}") for x in all_codes]