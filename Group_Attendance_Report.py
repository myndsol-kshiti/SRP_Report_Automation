
import pandas as pd
import numpy as np
import datetime
from datetime import datetime, date

Attendance_report = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\Shri Ram Pistons Reports\input\EmployeeData.xlsx",
                                  sheet_name="Employee Data Report")
Daily_Attendance_Report = pd.read_excel(r"C:\Users\kshiti.sinha\Desktop\projects\Shri Ram Pistons Reports\input\Daily_Attendance_Report.xlsx")
Daily_Attendance_Report['Shift'] = Daily_Attendance_Report['Shift Name'].astype(str)

Daily_Attendance_Report.loc[Daily_Attendance_Report['Shift'] == "06:00 M to 02:30 PM"] == "A"
# for i, row in Daily_Attendance_Report.iterrows():
#     if row['Shift'] == "06:00 M to 02:30 PM":
#         row['Shift'] = 'A'
#     elif row['Shift'] == "14:00 PM to 22:30 PM":
#         row['Shift'] = 'B'
#     elif row['Shift'] == "22:00 PM to 06:30 AM":
#         row['Shift'] = 'C'
#     elif row['Shift'] == "09:00 AM to 17:30 PM":
#         row['Shift'] = 'G'
#     else:
#         row['Shift'] = 'no shift'

Daily_Attendance_Report.to_excel('daily.xlsx')
Daily_manpower = pd.DataFrame(columns=['On Roll', 'A', 'G', 'B', 'C', 'Total'])

# Daily_manpower['Dept.'] = Attendance_report['DEPARTMENT'].drop_duplicates().dropna()

count = Attendance_report['DEPARTMENT'].value_counts()
Daily_manpower['On Roll'] = count
Daily_manpower.index.names = ['Dept.']


# Daily_manpower['On Roll'] = ['35','37','45','80']
# Daily_manpower['A'] = ['21','13','10','16']
# Daily_manpower['B'] = ['21','13','10','16']
# Daily_manpower['G'] = ['21','13','10','16']
# Daily_manpower['C'] = ['21','13','10','16']
Daily_manpower['Total'] = Daily_manpower['A'] + Daily_manpower['B'] + Daily_manpower['C'] + Daily_manpower['G']
# Daily_manpower = Daily_manpower['Dept.']
# Daily_manpower = Daily_manpower.drop_duplicates()
# Daily_manpower = Daily_manpower.dropna()
print(Daily_manpower)
# Daily_manpower.to_excel(r'Group_Attendance_Report.xlsx')
