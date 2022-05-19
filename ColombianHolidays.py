import datetime
from datetime import date
import holidays

def days_before_holidays():
    days_before_holiday = []
    for p, i in holidays.Colombia(years = 2022).items():
        day = p - datetime.timedelta(days=1)
        days_before_holiday.append(day)
    return days_before_holiday
