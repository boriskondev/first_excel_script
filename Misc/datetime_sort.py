from dateutil.parser import parse  # this library should be additionally installed

dates_list = ["27.08.1986", "26.08.1986", "28011900"]

for date in dates_list:
    if len(date) < 10:
        new_date = date[0:2] + "." + date[2:4] + "." + date[4:len(date)]
        dates_list.remove(date)
        dates_list.append(new_date)

print(dates_list)

[parse(x) for x in dates_list]

print(sorted(dates_list))