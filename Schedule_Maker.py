import datetime
import calendar
from xlsxwriter import Workbook


def schedule_creator(start, end):
    delta = end - start
    time = str(datetime.datetime.now().hour) + str(datetime.datetime.now().minute) + str(datetime.datetime.now().second)
    year_list = []
    month_list = []
    date_list = []
    year_key = []
    month_key = []
    date_dict = {}
    week_col = []
    schedule_book = Workbook(str(start.year)+str(start.month)+str(start.day)+"_"+time+'.xlsx')
    worksheet = schedule_book.add_worksheet('Schedule')
    cell_format = schedule_book.add_format({'bold':1,'align':'center','valign':'vcenter','border':1})
    blank_format = schedule_book.add_format()
    blank_format.set_left(7)
    blank_format.set_right(7)
    week_format = schedule_book.add_format()
    week_format.set_right(12)
    border_format = schedule_book.add_format()
    border_format.set_right(5)
    
    for i in range(0,delta.days):
        date = start + datetime.timedelta(days = i)
        if calendar.weekday(date.year,date.month,date.day) == 5 or calendar.weekday(date.year,date.month,date.day) == 6:
            continue        
        year_list.append(date.year)
        month_list.append(date.month)
        date_list.append(date)
    count = 0
    for i in year_list:
        if i not in date_dict.keys():
            date_dict[i] = {}
        if month_list[count] not in date_dict[i].keys():
            date_dict[i][month_list[count]] = []

        date_dict[i][month_list[count]].append(date_list[count])
        count += 1
    col = 1
    for year_key in date_dict.keys():
        worksheet.merge_range(0,col,0,col + year_list.count(year_key)-1,year_key,cell_format)
        for month_key in date_dict[year_key].keys():
            worksheet.merge_range(1,col,1,col + len(date_dict[year_key][month_key])-1,calendar.month_name[month_key],cell_format)        
            for i in date_dict[year_key][month_key]:
                worksheet.write(2,col,calendar.day_abbr[calendar.weekday(i.year,i.month,i.day)],cell_format)
                worksheet.write(3,col,i.day,cell_format)    
                if calendar.weekday(i.year,i.month,i.day) == 4:
                    week_col.append(col)
                col += 1
    worksheet.set_column(1,col,None,blank_format)
    for i in week_col:
        worksheet.set_column(i,i,None,week_format)
    worksheet.set_column(0,0,None,border_format)
    worksheet.set_column(col-1,col-1,None,border_format)
    worksheet.set_h_pagebreaks([79])
    worksheet.set_v_pagebreaks([35])
    schedule_book.close()



# Prompt User to input user start date 

   
while True:
    date_str = input('Enter the Start Date YYYY-MM-DD: \n (Leave Blank to Use Current Date)')
    if not date_str:
        start_date = datetime.datetime.now()
        break
    s_year, s_month, s_day = date_str.split('-')
    try:
        datetime.datetime(int(s_year), int(s_month), int(s_day))
    except ValueError:
        continue
    start_date = datetime.datetime(int(s_year),int(s_month),int(s_day))
    break
# Prompt User to input number of days
while True:
    end_date_str = input('Enter End Date YYYY-MM-DD: ')
    e_year, e_month, e_day = end_date_str.split('-')
    try:
        datetime.datetime(int(e_year), int(e_month), int(e_day))
    except ValueError:
        continue
    end_date = datetime.datetime(int(e_year),int(e_month),int(e_day))
    break

# Run Function
schedule_creator(start_date, end_date)