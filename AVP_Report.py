import pandas as pd
import numpy as np

old_attendance_report = pd.read_excel(
    r"C:\Users\kshiti.sinha\Desktop\projects\Shri Ram Pistons Reports\input\AttendanceReportWithTime.xls",
    sheet_name="Attendance Report")
avp_report = pd.DataFrame(
    columns=['Code', 'Name', 'Category', 'Line/Section', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
             '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28',
             '29', '30', '31', 'Grand Total'])
new_header = old_attendance_report.iloc[0]
old_attendance_report = old_attendance_report[1:]
old_attendance_report.columns = new_header
old_attendance_report['date_value'] = pd.to_datetime(old_attendance_report['Date']).dt.day.astype(str)
emp_value = old_attendance_report['Employee Code'].unique()
emp_value = emp_value

avp_report['Code'] = old_attendance_report['Employee Code'].drop_duplicates()
avp_report['Name'] = old_attendance_report['Employee Name'].drop_duplicates()
avp_report['Category'] = old_attendance_report['Employee Type'].drop_duplicates()
avp_report['Line/Section'] = old_attendance_report['Function'].drop_duplicates()
arrival = pd.to_datetime(old_attendance_report['Actual In Time'], format='%H:%M:%S')
departure = pd.to_datetime(old_attendance_report['Actual Out Time'], format='%H:%M:%S')
old_attendance_report['hours'] =abs((arrival - departure).astype('timedelta64[h]'))
print(emp_value)
for index,i in enumerate(emp_value):
    print(i)
    df = old_attendance_report[old_attendance_report['Employee Code'].str.contains((i))]
    monthly_hours = df['hours'].tolist()
    loop = len(monthly_hours)
    print(loop)
    if loop == 31:
        avp_report.at[index+1,'1'] = monthly_hours[0]
        avp_report.at[index+1,'2'] = monthly_hours[1]
        avp_report.at[index+1,'3'] = monthly_hours[2]
        avp_report.at[index + 1, '4'] = monthly_hours[3]
        avp_report.at[index + 1, '5'] = monthly_hours[4]
        avp_report.at[index + 1, '6'] = monthly_hours[5]
        avp_report.at[index + 1, '7'] = monthly_hours[6]
        avp_report.at[index + 1, '8'] = monthly_hours[7]
        avp_report.at[index + 1, '9'] = monthly_hours[8]
        avp_report.at[index + 1, '10'] = monthly_hours[9]
        avp_report.at[index + 1, '11'] = monthly_hours[10]
        avp_report.at[index + 1, '12'] = monthly_hours[11]
        avp_report.at[index + 1, '13'] = monthly_hours[12]
        avp_report.at[index + 1, '14'] = monthly_hours[13]
        avp_report.at[index + 1, '15'] = monthly_hours[14]
        avp_report.at[index + 1, '16'] = monthly_hours[15]
        avp_report.at[index + 1, '17'] = monthly_hours[16]
        avp_report.at[index + 1, '18'] = monthly_hours[17]
        avp_report.at[index + 1, '19'] = monthly_hours[18]
        avp_report.at[index + 1, '20'] = monthly_hours[19]
        avp_report.at[index + 1, '21'] = monthly_hours[20]
        avp_report.at[index + 1, '22'] = monthly_hours[21]
        avp_report.at[index + 1, '23'] = monthly_hours[22]
        avp_report.at[index + 1, '24'] = monthly_hours[23]
        avp_report.at[index + 1, '25'] = monthly_hours[24]
        avp_report.at[index + 1, '26'] = monthly_hours[25]
        avp_report.at[index + 1, '27'] = monthly_hours[26]
        avp_report.at[index + 1, '28'] = monthly_hours[27]
        avp_report.at[index + 1, '29'] = monthly_hours[28]
        avp_report.at[index + 1, '30'] = monthly_hours[29]
        avp_report.at[index + 1, '31'] = monthly_hours[30]
        avp_report.fillna(0)
        # avp_report['Grand Total'] = avp_report.sum(axis=1)
        avp_report['Grand Total'] = avp_report['1'] + avp_report['2'] + avp_report['3'] + avp_report['4'] + avp_report['5'] + avp_report['6'] + avp_report['7'] + avp_report['8'] +avp_report['9'] + avp_report['10'] + avp_report['11'] + avp_report['12'] + avp_report['13'] + avp_report['14'] + avp_report['15'] + avp_report['16'] + avp_report['17'] + avp_report['18'] + avp_report['19'] + avp_report['20'] + avp_report['21'] + avp_report['22'] + avp_report['23'] + avp_report['24'] + avp_report['25'] + avp_report['26'] + avp_report['27'] + avp_report['28'] + avp_report['29'] +avp_report['30'] + avp_report['31']

    elif loop == 30:
        avp_report.at[index + 1, '1'] = monthly_hours[0]
        avp_report.at[index + 1, '2'] = monthly_hours[1]
        avp_report.at[index + 1, '3'] = monthly_hours[2]
        avp_report.at[index + 1, '4'] = monthly_hours[3]
        avp_report.at[index + 1, '5'] = monthly_hours[4]
        avp_report.at[index + 1, '6'] = monthly_hours[5]
        avp_report.at[index + 1, '7'] = monthly_hours[6]
        avp_report.at[index + 1, '8'] = monthly_hours[7]
        avp_report.at[index + 1, '9'] = monthly_hours[8]
        avp_report.at[index + 1, '10'] = monthly_hours[9]
        avp_report.at[index + 1, '11'] = monthly_hours[10]
        avp_report.at[index + 1, '12'] = monthly_hours[11]
        avp_report.at[index + 1, '13'] = monthly_hours[12]
        avp_report.at[index + 1, '14'] = monthly_hours[13]
        avp_report.at[index + 1, '15'] = monthly_hours[14]
        avp_report.at[index + 1, '16'] = monthly_hours[15]
        avp_report.at[index + 1, '17'] = monthly_hours[16]
        avp_report.at[index + 1, '18'] = monthly_hours[17]
        avp_report.at[index + 1, '19'] = monthly_hours[18]
        avp_report.at[index + 1, '20'] = monthly_hours[19]
        avp_report.at[index + 1, '21'] = monthly_hours[20]
        avp_report.at[index + 1, '22'] = monthly_hours[21]
        avp_report.at[index + 1, '23'] = monthly_hours[22]
        avp_report.at[index + 1, '24'] = monthly_hours[23]
        avp_report.at[index + 1, '25'] = monthly_hours[24]
        avp_report.at[index + 1, '26'] = monthly_hours[25]
        avp_report.at[index + 1, '27'] = monthly_hours[26]
        avp_report.at[index + 1, '28'] = monthly_hours[27]
        avp_report.at[index + 1, '29'] = monthly_hours[28]
        avp_report.at[index + 1, '30'] = monthly_hours[29]
        avp_report.fillna(0)
        avp_report['Grand Total'] = avp_report['1'] + avp_report['2'] + avp_report['3'] + avp_report['4'] + avp_report['5'] + avp_report['6'] + avp_report['7'] + avp_report['8'] + avp_report['9'] + avp_report['10'] +  avp_report['11'] + avp_report['12'] + avp_report['13'] + avp_report['14'] + avp_report['15'] + avp_report['16'] + avp_report['17'] + avp_report['18'] + avp_report['19'] + avp_report['20'] + avp_report['21'] + avp_report['22'] + avp_report['23'] + avp_report['24'] + avp_report['25'] + avp_report['26'] + avp_report['27'] + avp_report['28'] + avp_report['29'] + avp_report['30']


        # avp_report['Grand Total'] = avp_report['1'] + avp_report['2']



    print(avp_report)
    avp_report.to_excel('AVP_Report.xlsx')
    print(df)





# for i,row in old_attendance_report.iterrows():
#         if row['date_value'] == 1:
#             row2['1'] == old_attendance_report['hours']
#             print(row2['1'])

# avp_report.to_excel('AVP_Report.xlsx')

# avp_report['Grand Total'] = old_attendance_report[col_list].sum(axis=1)



