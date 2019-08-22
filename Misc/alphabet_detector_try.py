from alphabet_detector import AlphabetDetector
ad = AlphabetDetector()

codes_to_remove = ["АВТ.КОД/АС:", "Авт. код: ", "Авт. код ", "Авт.код :", "Авт. код. ", "АВТ. КОД/ ", "АВТ.КОД/AC:",
                   "AUTH.CODE:", "AUTH CODE:", "Auth. Code ", "PRE-AUTH ", "ACC", "AC:", "АС:", "AC :", "АС :", "AC: ",
                   "AC", "АС", "AС", "АC", "AC ", "ac ", "Ас ", "ac", "Ac", "tid "]

status = {}


for code in codes_to_remove:
    is_Latin = ad.only_alphabet_chars(code, "LATIN")
    if is_Latin:
        status[code] = "YES"
    else:
        status[code] = "NO"

[print(k + " --> " + v) for k,v in sorted(status.items(), key=lambda kv: (kv[1]))]