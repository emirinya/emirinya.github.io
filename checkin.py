import time
import datetime
import calendar

week_day_dict = {
    0 : '周一',
    1 : '周二',
    2 : '周三',
    3 : '周四',
    4 : '周五',
    5 : '周六',
    6 : '周日'
}

def get_week_day(date):
    day = date.weekday()
    return week_day_dict[day]
 
print(get_week_day(datetime.datetime.now()))

mounth = calendar.monthcalendar(2019, 7)
mounthLine = []

for week in mounth:
    for date in week:
        if date != 0:
            mounthLine.append( [date,week_day_dict[week.index(date) % 7]] )

print(mounthLine)
