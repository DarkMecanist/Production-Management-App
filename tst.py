import datetime

def return_duration_between_dates(date_start, date_finish):
    duration = date_finish - date_start

    if duration.days == 2:
        return duration - datetime.timedelta(days=2)

    return duration

date_1 = datetime.datetime(year=2021, month=3, day=5, hour=23, minute=30)
date_2 = datetime.datetime(year=2021, month=3, day=8, hour=0, minute=30)
date_3 = datetime.datetime(year=2021, month=3, day=9, hour=0, minute=30)

duration = return_duration_between_dates(date_2, date_3)
print(duration)