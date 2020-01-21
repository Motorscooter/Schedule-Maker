import datetime
import calendar
from xlsxwriter import Workbook


def schedule_creator(year=0, month=0, day=0, num_of_days=0):
    day_list = calendar.weekheader(3)
    if year == 0 & month == 0 & day == 0:
        start_date = datetime.datetime.now()
    else:
        start_date = datetime.datetime(year, month, day)
    day_delta = datetime.timedelta(days=num_of_days)
    end_date = start_date + day_delta
    iter_date = start_date
    year_start_col = 1
    month_start_col = 1
    end_col = 1
    workbook = Workbook(str(start_date.year)+str(start_date.month) + str(start_date.day) + '.xlsx')
    worksheet = workbook.add_worksheet('Schedule')
    cell_format = workbook.add_format()
    cell_format.set_border(1)
    cell_format.set_bold()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    month_format = workbook.add_format()
    month_format.set_left(2)
    month_format.set_right(2)
    year_format = workbook.add_format()
    year_format.set_left(5)
    year_format.set_right(5)
    delta_days = end_date - iter_date
    while delta_days.days >= 0:
        while iter_date.month <= 12:
            max_days = max(calendar.monthrange(iter_date.year, iter_date.month))
            while iter_date.day <= max_days or delta_days.days >= 0:
                worksheet.write(2, end_col, day_list[calendar.weekday(iter_date.year, iter_date.month, iter_date.day)],
                                cell_format)
                worksheet.write(3, end_col, iter_date.day, cell_format)
                iter_date = iter_date + datetime.timedelta(days=1)
                end_col += 1
                delta_days = end_date - iter_date
            worksheet.merge_range(1, month_start_col, 1, end_col, calendar.month_name[iter_date.month], cell_format)
            worksheet.set_column(month_start_col, end_col, month_format)
            month_start_col = end_col + 1
        worksheet.merge_range(0,year_start_col, 0, end_col, cell_format)
        worksheet.set_column(year_start_col, end_col, year_format)
    workbook.close()


# Prompt User to input user start date
while True:
    date = input('Enter the Start Date YYYY-MM-DD: \n (Enter Nothing to Start from Current Date)')
    if not date:
        date = datetime.datetime.now()
        in_year = date.year
        in_month = date.month
        in_day = date.day
        break
    in_year, in_month, in_day = date.split('-')
    try:
        datetime.datetime(int(in_year), int(in_month), int(in_day))
    except ValueError:
        continue
    break
# Prompt User to input number of days
while True:
    number_of_days = input('Enter the number of days to create the schedule: ')
    try:
        int(number_of_days)
    except ValueError:
        continue
    number_of_days = int(number_of_days)
    break
# Run Function
schedule_creator(year=in_year, month=in_month, day=in_day, num_of_days=number_of_days)



