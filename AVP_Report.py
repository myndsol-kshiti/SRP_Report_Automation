import pandas as pd
import numpy as np

old_attendance_report = pd.read_excel(
    r"C:\Users\kshiti.sinha\Desktop\projects\Shri Ram Pistons Reports\input\AttendanceReportWithTime.xls",
    sheet_name="Attendance Report")
avp_report = pd.DataFrame(
    columns=['Code', 'Name', 'Category', 'Line/Section', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
             '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28',
             '29', '30', '31', 'Grand Total'])
col_list = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19',
            '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']
new_header = old_attendance_report.iloc[0]
old_attendance_report = old_attendance_report[1:]
old_attendance_report.columns = new_header
old_attendance_report['date_value'] = pd.to_datetime(old_attendance_report['Date']).dt.day.astype(str)
emp_value = old_attendance_report['Employee Code'].unique()
old_attendance_report

avp_report['Code'] = old_attendance_report['Employee Code'].drop_duplicates()
avp_report['Name'] = old_attendance_report['Employee Name'].drop_duplicates()
avp_report['Category'] = old_attendance_report['Employee Type'].drop_duplicates()
avp_report['Line/Section'] = old_attendance_report['Function'].drop_duplicates()
arrival = pd.to_datetime(old_attendance_report['Actual In Time'], format='%H:%M:%S')
departure = pd.to_datetime(old_attendance_report['Actual Out Time'], format='%H:%M:%S')
old_attendance_report['hours'] =abs((arrival - departure).astype('timedelta64[h]'))


for i,row in old_attendance_report.iterrows():
    for j,row2 in avp_report.iterrows():
        if row['date_value'] == 1:
            row2['1'] == old_attendance_report['hours']
            print(row2['1'])

print(avp_report)

avp_report.to_excel('AVP_Report.xlsx')

# avp_report['Grand Total'] = old_attendance_report[col_list].sum(axis=1)



