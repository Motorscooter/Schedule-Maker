import datetime
import Schedule_Maker

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
schedule_creator(year=int(in_year), month=int(in_month), day=int(in_day), num_of_days=number_of_days)