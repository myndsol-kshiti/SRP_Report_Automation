import pandas as pd
import numpy as np
import datetime
from datetime import datetime, date

daily_attendance_old = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\Shri Ram Pistons Reports\input\Daily_Attendance_Report.xlsx",
    sheet_name="Sheet1")

daily_attendance_new = pd.DataFrame(
    columns=['SR NO', 'CODE', 'NAME', 'DATE', 'SHIFT', 'ARRIVAL', 'DEPARTURE', 'WORK', 'Total', 'STATUS', 'PIR'])

# daily_attendance_new['SR NO'] = daily_attendance_old['']
daily_attendance_new['CODE'] = daily_attendance_old['Employee Code']
daily_attendance_new['NAME'] = daily_attendance_old['Employee Name']
daily_attendance_new['DATE'] = daily_attendance_old['Date'].dt.date.astype(str)

for i,row in daily_attendance_old.iterrows():
    if row['Shift Name'] == "06:00 AM to 02:30 PM":
        daily_attendance_new['SHIFT'] = "A"
    elif row['Shift Name'] == "14:00 PM to 22:30 PM":
        daily_attendance_new['SHIFT'] = "B"
    elif row['Shift Name'] == "22:00 PM to 06:30 AM":
        daily_attendance_new['SHIFT'] = "C"
    elif row['Shift Name'] == "09:00 AM to 17:30 PM":
        daily_attendance_new['SHIFT'] = "G"
    else:
        daily_attendance_new['SHIFT'] = "no shift"



daily_attendance_new['ARRIVAL'] = daily_attendance_old['In Time']
arrival = pd.to_datetime(daily_attendance_old['In Time'], format='%H:%M:%S')
#arrival = datetime.strptime(arrival_new,"%I:%M %p")

daily_attendance_new['DEPARTURE'] = daily_attendance_old['Out Time']
departure = pd.to_datetime(daily_attendance_old['Out Time'], format='%H:%M:%S')
daily_attendance_new['WORK'] = arrival - departure
daily_attendance_new['WORK'].dt.days.astype(float)

for i,row in daily_attendance_old.iterrows():
        if row['Actual Status'] == "Present":
            row['STATUS'] = "P"
        elif row['Actual Status'] == "Absent":
            row['STATUS'] = "A"
        else:
            row['STATUS'] = "CL for now"

for i,row in daily_attendance_new.iterrows():
    if row['WORK'] >= float(8.0):
        daily_attendance_new['PIR'] = "1.00"
    elif row['WORK'] <= float(4.0):
        daily_attendance_new['PIR'] = "0.50"
    else:
        daily_attendance_new['PIR'] = "0.00"


daily_attendance_new.replace(np.nan,"",regex=True)


# daily_attendance_new['WORK'] =

daily_attendance_new.to_excel('daily_attendance_report.xlsx',columns=['CODE', 'NAME', 'DATE', 'SHIFT', 'ARRIVAL', 'DEPARTURE', 'WORK', 'Total', 'STATUS', 'PIR'])