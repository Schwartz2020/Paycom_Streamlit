import pandas as pd
import numpy as np
import streamlit as st
import base64
from io import BytesIO
import datetime


def loadRawPaycomExcel(rawExcelFile):
    df = pd.read_excel(rawExcelFile, index_col=None,
                       usecols="A,B,C,G,H,I,J,K", engine="openpyxl", header=0)
    df = df.where(df.notnull(), None)
    df["Full_Name"] = df["Firstname"].map(str) + " " + df["Lastname"].map(str)
    df[["Date", "ClockIn"]] = df["InPunchTime"].str.split(' ', 1, expand=True)
    df[["Date2", "ClockOut"]] = df["OutPunchTime"].str.split(
        ' ', 1, expand=True)
    df["Work_Time_Frame"] = df["ClockIn"].map(
        str) + " - " + df["ClockOut"].map(str)
    df["Day"] = df.apply(lambda x: datetime.datetime.strptime(
        x["Date"], '%Y-%m-%d').strftime('%a'), axis=1)
    df = df.query('Department=="Driver"')
    rawData = df[['Full_Name', 'EECode', 'Day', 'Date',
                  'Work_Time_Frame', 'EarnHours', 'EarnCode', 'ClockIn', 'ClockOut']]
    rawData.rename(columns={"EarnHours": "Hours",
                            "EarnCode": "Sup_Info"}, inplace=True)
    return rawData

    # cleanRawDataDf = pd.DataFrame(dfList, columns=[
    #     'Full_Name', 'Work_code', 'Day', 'Date', 'Work_Time_Frame', 'Lunch', 'Hours', 'Sup_Info'])
    # return cleanRawDataDf


def getTeamWorkWeekStats(rawDataDf):
    rawDataNoPTODf = rawDataDf.query(
        '(Sup_Info != "PTO") & (Sup_Info != "TRN")')
    rawDataPTODf = rawDataDf.query('(Sup_Info == "PTO") | (Sup_Info == "TRN")')
    x = (rawDataNoPTODf.groupby(['Full_Name', 'Date']).sum().groupby(
        level=0).cumsum().rename(columns={'Hours': 'CumSum'})).reset_index()
    y = (rawDataNoPTODf.groupby(['Full_Name', 'Date']).sum()).reset_index()
    z = x.join(y, lsuffix='L_')
    z['Reg_Hours'] = np.where(z['CumSum'] <= 40, z['Hours'], np.where(
        z['CumSum']-z['Hours'] <= 40, z['Hours']-(z['CumSum']-40), 0))
    z['O_Hours'] = np.where(z['CumSum'] <= 40, 0, np.where(
        z['CumSum']-z['Hours'] <= 40, z['CumSum']-40, z['Hours']))
    # 2
    # teamWorkWeekStatsDf = rawDataDf.assign(
    #     ptoResult=np.where(
    #         rawDataDf['Sup_Info'] == 'PTO', rawDataDf.Hours, 0)
    # ).merge(z, how='left', left_on=['Full_Name', 'Date'], right_on=['Full_NameL_', 'DateL_'], suffixes=(None, "_y")).groupby(['Date', 'Day']).agg({'Reg_Hours': 'sum', 'O_Hours': 'sum', 'ptoResult': 'sum', 'Hours': 'sum'}).rename(columns={'Reg_Hours': 'Work_Hours', 'O_Hours': 'Overtime_Hours', 'ptoResult': 'PTO_Hours', 'Hours': 'Total_Hours'}).reset_index()
    appendedDf = z.append(rawDataPTODf, ignore_index=True)
    teamWorkWeekStatsDf = appendedDf.assign(
        ptoResult=np.where(
            appendedDf['Sup_Info'] == 'PTO', appendedDf.Hours, 0),
        trainResult=np.where(
            appendedDf['Sup_Info'] == 'TRN', appendedDf.Hours, 0)
    ).groupby(['Date']).agg({'Reg_Hours': 'sum', 'O_Hours': 'sum', 'ptoResult': 'sum', 'trainResult': 'sum', 'Hours': 'sum'}).rename(columns={'Reg_Hours': 'Work_Hours', 'O_Hours': 'Overtime_Hours', 'ptoResult': 'PTO_Hours', 'trainResult': 'Training_Hours', 'Hours': 'Total_Hours'}).reset_index()
    return teamWorkWeekStatsDf


def sumGreaterThanZero(pandasSeries):
    if pandasSeries.sum() == 1:
        return 'O'  # Worked and did not take lunch
    elif pandasSeries.sum() >= 1:
        return 'X'  # Worked and took lunch
    else:
        return None


def getDaysOfWeek(ungroupedData):
    ungroupedDataDaysOfWeek = ungroupedData.assign(
        Sunday=np.where(ungroupedData['Day'] == 'Sun', 1, 0),
        Monday=np.where(ungroupedData['Day'] == 'Mon', 1, 0),
        Tuesday=np.where(ungroupedData['Day'] == 'Tue', 1, 0),
        Wednesday=np.where(ungroupedData['Day'] == 'Wed', 1, 0),
        Thursday=np.where(ungroupedData['Day'] == 'Thu', 1, 0),
        Friday=np.where(ungroupedData['Day'] == 'Fri', 1, 0),
        Saturday=np.where(ungroupedData['Day'] == 'Sat', 1, 0)
    )
    groupedDataDaysOfWeek = ungroupedDataDaysOfWeek.groupby(['Full_Name']).agg(
        {'Sunday': sumGreaterThanZero, 'Monday': sumGreaterThanZero, 'Tuesday': sumGreaterThanZero, 'Wednesday': sumGreaterThanZero, 'Thursday': sumGreaterThanZero, 'Friday': sumGreaterThanZero, 'Saturday': sumGreaterThanZero})
    return groupedDataDaysOfWeek


def getNonOvertimeHours(pandasSeries):
    if pandasSeries.sum() > 40:
        return 40.0
    else:
        return pandasSeries.sum()


def getOvertimeHours(pandasSeries):
    if pandasSeries.sum() > 40:
        return pandasSeries.sum()-40.0
    else:
        return 0.0


def getDriverWorkDayStats(rawDataDf):
    rawDataNoPTODf = rawDataDf.query(
        '(Sup_Info != "PTO") & (Sup_Info != "TRN")')
    groupedDataDaysOfWeek = getDaysOfWeek(rawDataNoPTODf)
    rdnpGroupByDf = rawDataNoPTODf.groupby(['Full_Name'])
    daysWorked = rdnpGroupByDf['Date'].nunique().to_frame(
        name='Days_Worked')
    driverWorkDayStatsDf = daysWorked.join(
        rdnpGroupByDf.agg({'Hours': getNonOvertimeHours}).rename(columns={'Hours': 'Work_Hours'})).join(rdnpGroupByDf.agg({'Hours': getOvertimeHours}).rename(columns={'Hours': 'Overtime_Hours'})).join(rdnpGroupByDf.agg({'Hours': 'sum'}).rename(columns={'Hours': 'Total_Work_Hours'})).join(groupedDataDaysOfWeek).reset_index()
    driverWorkDayStatsDf['Hours_Owed'] = np.where(
        (8*driverWorkDayStatsDf['Days_Worked']-driverWorkDayStatsDf['Work_Hours']) > 0, 8*driverWorkDayStatsDf['Days_Worked']-driverWorkDayStatsDf['Work_Hours'], 0)
    return driverWorkDayStatsDf


def getPTOStats(rawDataDf):
    rawDataPTODf = rawDataDf.query('Sup_Info=="PTO"')
    groupedDataDaysOfWeek = getDaysOfWeek(rawDataPTODf)
    rdpGroupByDf = rawDataPTODf.groupby(['Full_Name'])
    daysPTO = rdpGroupByDf['Date'].nunique().to_frame(name='Days_PTO')
    ptoStatsDf = daysPTO.join(rdpGroupByDf.agg({'Sup_Info': 'count'}).join(rdpGroupByDf.agg({'Hours': 'sum'})).join(groupedDataDaysOfWeek).rename(columns={
        'Sup_Info': 'PTO_Instances', 'Hours': 'PTO_Hours'})).reset_index()
    return ptoStatsDf


def getTrainingStats(rawDataDf):
    rawDataTrainDf = rawDataDf.query('Sup_Info=="TRN"')
    groupedDataDaysOfWeek = getDaysOfWeek(rawDataTrainDf)
    rdpGroupByDf = rawDataTrainDf.groupby(['Full_Name'])
    daysTrain = rdpGroupByDf['Date'].nunique().to_frame(name='Days_Train')
    trainingStatsDf = daysTrain.join(rdpGroupByDf.agg({'Sup_Info': 'count'}).join(rdpGroupByDf.agg({'Hours': 'sum'})).join(groupedDataDaysOfWeek).rename(columns={
        'Sup_Info': 'Train_Instances', 'Hours': 'Train_Hours'})).reset_index()
    return trainingStatsDf


def timeDiffCalculation(clockIn, clockOut):
    timeA = datetime.datetime.strptime(clockIn, "%I:%M %p")
    timeB = datetime.datetime.strptime(clockOut, "%I:%M %p")
    timeDiff = ((timeB-timeA).total_seconds()/60)
    return timeDiff


def timeDiffCalculationModified(clockIn, clockOut):
    try:
        timeA = datetime.datetime.strptime(clockIn, "%I:%M %p")
        timeB = datetime.datetime.strptime(clockOut, "%I:%M %p")
        timeDiff = ((timeB-timeA).total_seconds()/60)
        return timeDiff
    except Exception:
        return -1  # To negate 5-hr condition


def getBreakStats(rawDataDf):
    # Only work instances
    rawDataNoPTODf = rawDataDf.query(
        '(Sup_Info != "PTO") & (Sup_Info != "TRN")')
    rawDataNoPTODf["Mins_Worked"] = rawDataNoPTODf.apply(
        lambda row: timeDiffCalculation(row.ClockIn, row.ClockOut), axis=1)
    # Create lag
    shifted = rawDataNoPTODf.groupby(["Full_Name", "Date"]).shift(-1)
    rawDataNoPTODf = rawDataNoPTODf.join(
        shifted.rename(columns=lambda x: x+"_lag"))
    #
    groupedData = rawDataNoPTODf.groupby(['Full_Name', 'Date'])
    groupCounts = groupedData.size().to_frame(name='Breaks')
    groupCounts["Breaks"] = groupCounts["Breaks"]-1
    #
    cleanerBreakStatsDf = groupCounts.join(groupedData.agg({'ClockIn': 'first'})).join(groupedData.agg({'ClockIn_lag': 'first'}).rename(columns={'ClockIn_lag': 'Second_ClockIn'})).join(groupedData.agg({'ClockOut': 'first'}).rename(columns={'ClockOut': 'First_ClockOut'})).join(groupedData.agg(
        {'ClockOut': 'last'})).join(groupedData.agg({'Mins_Worked': 'sum'})).join(groupedData.agg({'Hours': 'sum'}).rename(columns={'Hours': 'Hours_Worked'})).reset_index()
    cleanerBreakStatsDf["Mins_On_Break"] = cleanerBreakStatsDf.apply(
        lambda x: timeDiffCalculation(x.ClockIn, x.ClockOut)-x.Mins_Worked, axis=1)
    cleanerBreakStatsDf["Break_Check"] = cleanerBreakStatsDf.apply(
        lambda x: timeDiffCalculationModified(x.First_ClockOut, x.Second_ClockIn), axis=1)
    cleanerBreakStatsDf["5_hr_Compliant"] = cleanerBreakStatsDf.apply(
        lambda x: "Yes" if (timeDiffCalculation(x.ClockIn, x.First_ClockOut) < 300) & (x.Break_Check >= 30) & (x.Hours_Worked >= 7.5) else ("N/A" if (x.Hours_Worked < 7.5) else "No"), axis=1)
    breakStatsDf = cleanerBreakStatsDf[[
        "Full_Name", "Date", "Hours_Worked", "Breaks", "Mins_On_Break", "ClockIn", "First_ClockOut", "Second_ClockIn", "5_hr_Compliant"]]
    return breakStatsDf


def getMissingLunchInstances(rawDataDf):
    rawDataNoPTODf = rawDataDf.query(
        '(Sup_Info != "PTO") & (Sup_Info != "TRN")')
    missingLunchInstancesDf = rawDataNoPTODf.groupby(['Full_Name', 'Date']).filter(
        lambda df: len(df) < 2)
    return missingLunchInstancesDf


def getNotClockedOutInstances(rawDataDf):
    notClockedOutInstancesDf = rawDataDf.query(
        'Work_Time_Frame.str.count(":") < 2')
    return notClockedOutInstancesDf


def sendDataToExcelFile(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf,
                        ptoStatsDf,  trainingStatsDf, breakStatsDf, missingLunchInstancesDf, notClockedOutInstancesDf):

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    cell_format = workbook.add_format()
    cell_format.set_align('center')
    #
    rawDataDf.to_excel(writer, index=False, startrow=1,
                       header=False, sheet_name='Raw Data')
    columns1 = [{'header': column} for column in rawDataDf.columns]
    (max_row1, max_col1) = rawDataDf.shape
    (writer.sheets['Raw Data']).add_table(
        0, 0, max_row1, max_col1-1, {'columns': columns1})
    (writer.sheets['Raw Data']).set_column('A:M', 18, cell_format)
    #
    teamWorkWeekStatsDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Team Hours')
    columns2 = [{'header': column} for column in teamWorkWeekStatsDf.columns]
    (max_row2, max_col2) = teamWorkWeekStatsDf.shape
    (writer.sheets['Team Hours']).add_table(
        0, 0, max_row2, max_col2-1, {'columns': columns2})
    (writer.sheets['Team Hours']).set_column('A:M', 18, cell_format)
    #
    driverWorkDayStatsDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Driver Hours')
    columns3 = [{'header': column} for column in driverWorkDayStatsDf.columns]
    (max_row3, max_col3) = driverWorkDayStatsDf.shape
    (writer.sheets['Driver Hours']).add_table(
        0, 0, max_row3, max_col3-1, {'columns': columns3})
    (writer.sheets['Driver Hours']).set_column('A:M', 14, cell_format)
    #
    ptoStatsDf.to_excel(writer, index=False, startrow=1,
                        header=False, sheet_name='PTO Hours')
    columns4 = [{'header': column} for column in ptoStatsDf.columns]
    (max_row4, max_col4) = ptoStatsDf.shape
    (writer.sheets['PTO Hours']).add_table(
        0, 0, max_row4, max_col4-1, {'columns': columns4})
    (writer.sheets['PTO Hours']).set_column('A:M', 18, cell_format)
    #
    trainingStatsDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Training Stats')
    columns5 = [{'header': column}
                for column in trainingStatsDf.columns]
    (max_row5, max_col5) = trainingStatsDf.shape
    (writer.sheets['Training Stats']).add_table(
        0, 0, max_row5, max_col5-1, {'columns': columns5})
    (writer.sheets['Training Stats']
     ).set_column('A:M', 18, cell_format)
    #
    breakStatsDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Break Stats')
    columns6 = [{'header': column}
                for column in breakStatsDf.columns]
    (max_row6, max_col6) = breakStatsDf.shape
    (writer.sheets['Break Stats']).add_table(
        0, 0, max_row6, max_col6-1, {'columns': columns6})
    (writer.sheets['Break Stats']
     ).set_column('A:M', 18, cell_format)
    #
    missingLunchInstancesDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Missing Lunches')
    columns7 = [{'header': column}
                for column in missingLunchInstancesDf.columns]
    (max_row7, max_col7) = missingLunchInstancesDf.shape
    (writer.sheets['Missing Lunches']).add_table(
        0, 0, max_row7, max_col7-1, {'columns': columns7})
    (writer.sheets['Missing Lunches']).set_column('A:M', 18, cell_format)
    #
    notClockedOutInstancesDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Missing Clockouts')
    columns8 = [{'header': column}
                for column in notClockedOutInstancesDf.columns]
    (max_row8, max_col8) = notClockedOutInstancesDf.shape
    (writer.sheets['Missing Clockouts']).add_table(
        0, 0, max_row8, max_col8-1, {'columns': columns8})
    (writer.sheets['Missing Clockouts']).set_column('A:M', 18, cell_format)
    #
    writer.save()
    processed_data = output.getvalue()
    return processed_data


def get_table_download_link(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf,
                            ptoStatsDf,  trainingStatsDf, breakStatsDf, missingLunchInstancesDf, notClockedOutInstancesDf):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = sendDataToExcelFile(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf, ptoStatsDf,
                              trainingStatsDf, breakStatsDf, missingLunchInstancesDf, notClockedOutInstancesDf)
    b64 = base64.b64encode(val)  # val looks like b'...'
    # decode b'abc' => abc
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="Paycom_Summary.xlsx">Download Paycom Summary Excel File</a>'


st.set_page_config(page_title='Paycom Summary Generator')
st.title("Paycom Summary Generator")
st.subheader("Please upload your Paycom Excel File:")

uploaded_file = st.file_uploader(
    'Select the Paycom Time Sheet .xlsx file', type='xlsx')
if uploaded_file:
    st.write('File Uploaded.')
    rawDataDf = loadRawPaycomExcel(uploaded_file)
    teamWorkWeekStatsDf = getTeamWorkWeekStats(rawDataDf)
    driverWorkDayStatsDf = getDriverWorkDayStats(rawDataDf)
    ptoStatsDf = getPTOStats(rawDataDf)
    trainingStatsDf = getTrainingStats(rawDataDf)
    breakStatsDf = getBreakStats(rawDataDf)
    missingLunchInstancesDf = getMissingLunchInstances(rawDataDf)
    notClockedOutInstancesDf = getNotClockedOutInstances(rawDataDf)
    # Send out data
    st.markdown(get_table_download_link(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf, ptoStatsDf,
                                        trainingStatsDf, breakStatsDf, missingLunchInstancesDf, notClockedOutInstancesDf), unsafe_allow_html=True)
    st.write("File has been processed. Click link to download file.")
