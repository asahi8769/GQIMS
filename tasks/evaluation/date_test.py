from datetime import datetime, timedelta
from pytimekr import pytimekr

'20220203'
start = datetime.strptime("20220131", "%Y%m%d").date()
end = datetime.strptime("20220220", "%Y%m%d").date()


def count_working_days(start, end, *date):

    year_start = start.year
    year_end = end.year
    days_list = [start + timedelta(days=x) for x in range((end - start).days)]
    print(days_list)
    holidays = pytimekr.holidays(year=year_start) + pytimekr.holidays(year=year_end) + [datetime.strptime(i, "%Y%m%d").date() for i in date]
    print(holidays)
    dates_list = [i for i in days_list if i.weekday() not in [5, 6] and i not in holidays]
    print([i for i in days_list if i.weekday() not in [5, 6] and i not in holidays])
    return len(dates_list)


print(count_working_days(start, end))

"""
20220131 - 20220203
[datetime.date(2022, 1, 31), datetime.date(2022, 2, 1), datetime.date(2022, 2, 2)]


"""
