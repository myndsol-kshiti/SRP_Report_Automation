
import pandas as pd
import numpy as np
import datetime
from datetime import datetime, date

daily_attendance_old = pd.read_excel(r"C:\Users\hrithik.chauhan\Downloads\Daily_Attendance_Report.xlsx",
                                     sheet_name="Sheet1")

daily_attendance_new = pd.DataFrame(
    columns=['SR NO', 'CODE', 'NAME', 'DATE', 'SHIFT', 'ARRIVAL', 'DEPARTURE', 'WORK', 'Total', 'STATUS', 'PIR'])

# daily_attendance_new['SR NO'] = daily_attendance_old['']
daily_attendance_new['CODE'] = daily_attendance_old['Employee Code']
daily_attendance_new['NAME'] = daily_attendance_old['Employee Name']
daily_attendance_new['Shift_Name'] = daily_attendance_old['Shift Name']

daily_attendance_new['DATE'] = daily_attendance_old['Date'].dt.date.astype(str)

for i, row in daily_attendance_new.iterrows():
    if row['Shift_Name'] == "06:00 AM to 02:30 PM":

        row['SHIFT'] = "A"
    elif row['Shift_Name'] == "14:00 PM to 22:30 PM":
        row['SHIFT'] = "B"
    elif row['Shift_Name'] == "22:00 PM to 06:30 AM":
        row['SHIFT'] = "C"
    elif row['Shift_Name'] == "09:00 AM to 17:30 PM":
        row['SHIFT'] = "G"
    else:
        row['SHIFT'] = "no shift"


daily_attendance_new['ARRIVAL'] = daily_attendance_old['In Time']
arrival = pd.to_datetime(daily_attendance_old['In Time'], format='%H:%M:%S')
#arrival = datetime.strptime(arrival_new,"%I:%M %p")

daily_attendance_new['DEPARTURE'] = daily_attendance_old['Out Time']
departure = pd.to_datetime(daily_attendance_old['Out Time'], format='%H:%M:%S')
daily_attendance_new['WORK'] = abs((arrival - departure).astype('timedelta64[h]'))



daily_attendance_new['Total'] = "0.0"
daily_attendance_new['STATUS'] = daily_attendance_old['Actual Status']

count = 0;
for i, row in daily_attendance_new.iterrows():

    if row['WORK'] >= float(8.0):
        temp = i + 2
        count = count + 1
        daily_attendance_new.loc[i, 'PIR'] = "1.00"
    elif row['WORK'] == float(4.0):
        temp = i + 2
        count = count + 1
        daily_attendance_new.loc[i, 'PIR'] = "0.50"
    else:
        daily_attendance_new.loc[i, 'PIR'] = "0.00"

daily_attendance_new.replace(np.nan,"",regex=True)


daily_attendance_new.to_excel(r'C:\Users\hrithik.chauhan\Desktop\format.xlsx', columns=['SR NO','CODE'	,'NAME'	,'DATE'	,'SHIFT','ARRIVAL','DEPARTURE','WORK'  ,'Total' ,'STATUS','PIR'])
