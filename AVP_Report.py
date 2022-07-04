import pandas as pd
import numpy as np

old_attendance_report = pd.read_excel(
    r"C:\Users\kshiti.sinha\Desktop\projects\Shri Ram Pistons Reports\input\AttendanceReportWithTime.xls",
    sheet_name="Attendance Report")
avp_report = pd.DataFrame(
    columns=['Code', 'Name', 'Category', 'Line/Section', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
             '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28',
             '29', '30', '31', 'Grand Total'])
print(old_attendance_report)
col_list = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19',
            '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']
new_header = old_attendance_report.iloc[0]
old_attendance_report = old_attendance_report[1:]
old_attendance_report.columns = new_header
avp_report['date_value'] = pd.to_datetime(old_attendance_report['Date']).dt.day.astype(str)

# for i,row in avp_report.iterrows():
#     if row['date_value'] ==

# for i,row in old_attendance_report.iterrows():
#     print(row['Date'])
# avp_report.index = [x for x in range(1, len(avp_report.values)+1)]
avp_report['Code'] = old_attendance_report['Employee Code']
avp_report['Name'] = old_attendance_report['Employee Name']
avp_report['Category'] = old_attendance_report['Employee Type']
avp_report['Line/Section'] = old_attendance_report['Function']
# len(avp_report)
# avp_report = avp_report.drop_duplicates(subset=['Code','Name','Category','Function'])
# len(avp_report)

avp_report.to_excel('AVP_Report.xlsx')

# avp_report['Grand Total'] = old_attendance_report[col_list].sum(axis=1)

# print(avp_report)

# print(avp_report)


