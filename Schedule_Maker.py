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
    cell_format = workbook.add_format({'bold':1,'align':'center','valign':'vcenter','border':1})
    month_format = workbook.add_format({'bold':1,'align':'center','valign':'vcenter','left':1,'right':1})
    year_format = workbook.add_format({'bold':1,'align':'center','valign':'vcenter','left':1,'right':1})
    delta_days = end_date - iter_date
    while delta_days.days >= 0:
        while iter_date.month <= 12:
            max_days = max(calendar.monthrange(iter_date.year, iter_date.month))
            while True:
                if calendar.weekday(iter_date.year,iter_date.month, iter_date.day) > 4:
                    iter_date = iter_date + datetime.timedelta(days=1)
                    delta_days = end_date - iter_date
                    if iter_date.day == max_days or delta_days.days == 0:
                        end_col += 1
                        break
                    continue
                if iter_date.day == max_days or delta_days.days == 0:
                    worksheet.write(2, end_col,
                                    calendar.day_abbr[calendar.weekday(iter_date.year, iter_date.month, iter_date.day)],
                                    cell_format)
                    worksheet.write(3, end_col, iter_date.day, cell_format)
                    end_col += 1
                    break
                worksheet.write(2, end_col, calendar.day_abbr[calendar.weekday(iter_date.year, iter_date.month, iter_date.day)],
                                cell_format)
                worksheet.write(3, end_col, iter_date.day, cell_format)
                iter_date = iter_date + datetime.timedelta(days=1)
                end_col += 1
                delta_days = end_date - iter_date
            worksheet.merge_range(1, month_start_col, 1, end_col, calendar.month_name[iter_date.month], cell_format)
            month_start_col = end_col

            if delta_days.days == 0 or iter_date.month == 12:
                break
            iter_date = iter_date + datetime.timedelta(days=1)
        worksheet.merge_range(0,year_start_col, 0, end_col,iter_date.year, cell_format)
        worksheet.set_column(month_start_col, end_col, None, year_format)
        if delta_days.days == 0:
            break
        iter_date = iter_date + datetime.timedelta(days=1)
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



