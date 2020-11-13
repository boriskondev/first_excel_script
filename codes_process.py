import docx

new_codes = ['АВТ. КОД :', 'Auth. Code:', 'Auth. Code:', 'AC,:', 'AUTH. CODE :', 'АВТ.КОД/АС :', 'АВТ.КОД/АС :', 'Авт.код /AC', 'Auth. Code:', 'Авт. Код :', 'Авт. Код :', 'AUTH. CODE :']

project_config = docx.Document("config_new.docx")
codes_to_remove = project_config.paragraphs[3].text.split("=")[1].strip().replace("\"", "").split(", ")

all_codes = list(set(new_codes + codes_to_remove))

all_codes.sort()
all_codes.sort(key=len, reverse=True)

print(all_codes)