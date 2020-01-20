import numpy as np
import datetime
import calendar
from xlsxwriter import Workbook
cont_flag = False

# Prompt User to input user start date
while cont_flag:
    date = input('Enter the Start Date YYYY/MM/DD: \n (Enter Nothing to Start from Current Date)')
    if not date:
        break
    in_year, in_month, in_day = date.split('/')
    try:
        datetime.datetime(int(in_year), int(in_month), int(in_day))
    except ValueError:
        continue
    cont_flag = True
cont_flag = False
# Prompt User to input number of days
while cont_flag:
    number_of_days = input('Enter the number of days to create the schedule: ')
    try:
        int(number_of_days)
    except ValueError:
        continue
    cont_flag = True
# Run Function


def schedule_creator(year=0, month=0, day=0, num_of_days=0):
    day_list = calendar.weekheader(3)
    if year == 0 & month == 0 & day == 0:
        start_date = datetime.datetime.now()
    else:
        start_date = datetime.datetime(year, month, day)
    day_delta = datetime.timedelta(days=number_of_days)
    end_date = start_date + day_delta
    iter_date = start_date
    start_col = 1
    end_col = 1
    workbook = Workbook(str(start_date.year)+str(start_date.month) + str(start_date.day) + '.xlsx')
    worksheet = workbook.add_worksheet('Schedule')
    cell_format = workbook.add_format({'bold': True, 'valign': True})
    while end_date - iter_date <= 0:
        while iter_date.month <= 12:
            while iter_date.day <= max(calendar.monthrange(iter_date.year, iter_date.month)):
                worksheet.write(2, end_col, day_list[calendar.weekday(iter_date.year,iter_date.month, iter_date.day)],
                                cell_format)
                worksheet.write(3, end_col, iter_date.day, cell_format)
                iter_date = iter_date + datetime.timedelta(days=1)
                end_col += 1
            worksheet.merge_range(1, start_col, 1, end_date, calendar.month_name[iter_date.month], cell_format)