import pandas as pd
import numpy as np

old_attendance_report = pd.read_excel(
    r"C:\Users\kshiti.sinha\Desktop\projects\Shri Ram Pistons Reports\input\AttendanceReportWithTime.xls",
    sheet_name="Attendance Report")
avp_report = pd.DataFrame(
    columns=['Code', 'Name', 'Category', 'Line/Section', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
             '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28',
             '29', '30', '31', 'Grand Total'])
col_list = ['1', '2', '3']
new_header = old_attendance_report.iloc[0]
old_attendance_report = old_attendance_report[1:]
old_attendance_report.columns = new_header
old_attendance_report['date_value'] = pd.to_datetime(old_attendance_report['Date']).dt.day.astype(str)
emp_value = old_attendance_report['Employee Code'].unique()

avp_report['Code'] = old_attendance_report['Employee Code'].drop_duplicates()
avp_report['Name'] = old_attendance_report['Employee Name'].drop_duplicates()
avp_report['Category'] = old_attendance_report['Employee Type'].drop_duplicates()
avp_report['Line/Section'] = old_attendance_report['Function'].drop_duplicates()
arrival = pd.to_datetime(old_attendance_report['Actual In Time'], format='%H:%M:%S')
departure = pd.to_datetime(old_attendance_report['Actual Out Time'], format='%H:%M:%S')
old_attendance_report['hours'] =abs((arrival - departure).astype('timedelta64[h]'))


for i in emp_value:
    print(i)
    df = old_attendance_report[old_attendance_report['Employee Code'].str.contains((i))]
    monthly_hours = df['hours'].tolist()
    for count in range(0,len(df)):
        avp_report.at[count,'1'] = monthly_hours[0]
        avp_report.at[count,'2'] = monthly_hours[1]
        avp_report.at[count,'3'] = monthly_hours[2]
    # print(avp_report)
    avp_report.to_excel('test_avp.xlsx')
    print(df)





# for i,row in old_attendance_report.iterrows():
#         if row['date_value'] == 1:
#             row2['1'] == old_attendance_report['hours']
#             print(row2['1'])

# avp_report.to_excel('AVP_Report.xlsx')

# avp_report['Grand Total'] = old_attendance_report[col_list].sum(axis=1)



