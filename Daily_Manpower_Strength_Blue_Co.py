import pandas as pd
import numpy as np
import datetime
from datetime import datetime, date

Attendance_report = pd.read_excel(r"C:\Users\hrithik.chauhan\Downloads\EmployeeData.xlsx",
                                  sheet_name="Employee Data Report")

Daily_manpower = pd.DataFrame(columns=['DEPARTMENT1', 'On Roll', 'A', 'G', 'B', 'C', 'Total'])

daily_attendance_old = pd.read_excel(r"C:\Users\hrithik.chauhan\Downloads\Daily_Attendance_Report.xlsx",
                                     sheet_name="Sheet1")
# Daily_manpower['Dept.'] = Attendance_report['DEPARTMENT'].drop_duplicates()
# Daily_manpower['dept1'] = " "

Daily_manpower['A'] = " "
Daily_manpower['G'] = " "
Daily_manpower['B'] = " "
Daily_manpower['C'] = " "
Daily_manpower['Total'] = " "
count = Attendance_report['DEPARTMENT'].value_counts()
Daily_manpower['On Roll'] = count
Daily_manpower['DEPARTMENT'] = Daily_manpower.index
Daily_manpower.index.names = ['DEPARTMENT1']

Daily_manpower.to_excel(r'C:\Users\hrithik.chauhan\Desktop\daily.xlsx')
# Daily_manpower['dept1'] = Daily_manpower['Dept.']
df = daily_attendance_old
df.groupby(["DEPARTMENT", "Shift Duration"]).size()
ab = df.groupby(["DEPARTMENT", "Shift Duration"]).size().reset_index(name="Time")
ab.rename(columns={'Shift Duration': 'Shift_Duration'}, inplace=True)
print(ab)
ab.to_excel(r"C:\Users\hrithik.chauhan\Desktop\Book1.xlsx")
#  df.groupby(["Group", "Size"]).size().reset_index(name="Time")
# df.groupby(["Group", "Size"]).size().reset_index(name="Time")
inner_join = pd.merge(Daily_manpower,
                      ab,
                      on='DEPARTMENT',
                      how='left')

inner_join.to_excel(r"C:\Users\hrithik.chauhan\Desktop\demo.xlsx")
inner_join
for j, row in Daily_manpower.iterrows():

    for i, row1 in inner_join.iterrows():

        if row1['DEPARTMENT'] == row['DEPARTMENT']:
            if row1['Shift_Duration'] == "06:00am-02:30pm":
                row['A'] = row1['Time']
                Daily_manpower.loc[j, 'A'] = row1['Time']
            elif row1['Shift_Duration'] == "02:00pm-10:30pm":
                row['B'] = row1['Time']
                Daily_manpower.loc[j, 'B'] = row1['Time']
            elif row1['Shift_Duration'] == "10:00pm-06:30am":
                row['C'] = row1['Time']
                Daily_manpower.loc[j, 'C'] = row1['Time']
            elif row1['Shift_Duration'] == "09:00am-05:30pm":
                row['G'] = row1['Time']
                Daily_manpower.loc[j, 'G'] = row1['Time']
            print("1st loop", row['DEPARTMENT'], row['A'], row['B'])

print(Daily_manpower)
column_names = ['A', 'B', 'G', 'C']
Daily_manpower['Total'] = Daily_manpower[column_names].sum(axis=1)
Daily_manpower.to_excel(r'C:\Users\hrithik.chauhan\Desktop\daily.xlsx',
                        columns=['DEPARTMENT', 'On Roll', 'A', 'G', 'B', 'C', 'Total'])
