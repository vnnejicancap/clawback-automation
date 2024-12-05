from io import TextIOWrapper
import os
import openpyxl
import openpyxl.utils
import openpyxl.utils.dataframe
import pandas as pd
import numpy as np
import datetime as datetime
from sqlalchemy import create_engine
from sqlalchemy import text
from pathlib import Path
import glob
import urllib
import sys
import shutil
from dateutil.parser import parse
import re
import warnings
import time
from bs4 import BeautifulSoup
import PySimpleGUI as sg
import holidays
#import inspect
import json
import getpass
from babel.dates import format_date
import win32com.client as client
import re


#progName = inspect.getframeinfo(inspect.currentframe()).filename
#progPath  = os.path.dirname(os.path.abspath(progName))

exe_path = sys.executable
# Extract the directory part of the path
progPath = os.path.dirname(exe_path)
currentUser = getpass.getuser()
# progPath = "C:/Apps/Clawback/"

with open(progPath+"/config.json") as json_data:
    data = json.load(json_data)
ClawbackTestFolder = data["ClawbackTestFolder"]
ClawbackFiles = data["ClawbackFiles"]
ClawbackFiles = ClawbackFiles.replace("currentUser", currentUser)
ClawbackOutFolder = data["ClawbackOutFolder"]
ClawbackLog = data["ClawbackLog"]

currentYear = datetime.date.today().year
canadaHolidays = list((dict(holidays.Canada(years = currentYear, subdiv = "ON").items())).keys())

def connectionEngine():
    """Connection engine for connecting to the CIBC Database on PRDEDW001 server using Windows Authentication"""
    # Define the connection string using Windows Authentication
    conn_str = (
        r'Driver={ODBC Driver 17 for SQL Server};'\
        r"Server=PRDEDW001;"\
        r"Database=CIBC;"\
        r"Trusted_Connection=yes;"
    )
    # Connect to the SQL Server
    quoted = urllib.parse.quote_plus(conn_str)
    engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted), future = True)
    return engine

def connectionTestEngine():
    """Connection engine for connecting to the Test Database on TORSQL001 server using Windows Authentication"""
    # Define the connection string using Windows Authentication
    conn_str = (
        r'Driver={ODBC Driver 17 for SQL Server};'\
        r"Server=TORSQL001;"\
        r"Database=Test;"\
        r"Trusted_Connection=yes;"
    )
    # Connect to the SQL Server
    quoted = urllib.parse.quote_plus(conn_str)
    engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))
    return engine

def get_last_working_date(dt:datetime.datetime)->datetime.datetime:
    """Returns the last past working date from a given date. Runs recursively."""
    if dt.date().weekday() > 4 or dt.date() in canadaHolidays:
        dt = dt-datetime.timedelta(days=1)
        return get_last_working_date(dt)
    else:
        return dt

def date_to_locale(dateIn:datetime.datetime, province='ON'):
    """
    The function `date_to_locale` converts a given date to a locale-specific format based on the
    province provided.
    
    :param dateIn: The `dateIn` parameter is a datetime object representing a specific date and time
    :type dateIn: datetime.datetime
    :param province: The `province` parameter is used to specify the province for which the date formatting should be applied.
    If the province is 'QC' (Quebec), the date is formatted in French, otherwise, defaults to ON (optional)
    :return: The function `date_to_locale` returns a formatted date based on the input date and  province. 
    If the province is 'QC' (Quebec), the date is formatted in French using the format "EEEE, d MMMM, yyyy". 
    Otherwise, the date is formatted in English using the 'long' format.
    """
    if province == 'QC':
        dateOut = format_date(dateIn, locale='fr_FR', format="EEEE, d MMMM, yyyy")
    else:
        dateOut = format_date(dateIn, locale='en', format="EEEE, MMMM d, yyyy")
    return dateOut


def start_logger():
    """Initates logging of the clawback process to a logfile"""
    logTimeFormat = '%Y_%m_%d_%H_%M_%S'
    outputFormat = '%Y-%m-%d %H:%M:%S'
    logFile= open(ClawbackLog+"/logger_clawback_"+str(datetime.datetime.now().strftime(logTimeFormat))+".txt","w+")
    logMessage = "Started Clawback Process"
    print(str(datetime.datetime.now().strftime(outputFormat)) +" "+logMessage)
    logFile.write(str(datetime.datetime.now().strftime(outputFormat)) +" "+logMessage+'\n')
    return logFile

def write_to_logger(logMessage):
    """Write to the open log file\n
    param logMesage: type(str) Text to be written to log"""
    outputFormat = '%Y-%m-%d %H:%M:%S'
    print(str(datetime.datetime.now().strftime(outputFormat)) +" "+ logMessage)
    logFile.write(str(datetime.datetime.now().strftime(outputFormat)) +" "+logMessage+'\n')

def raise_fn_exception(fnName, err, lineNo, errMsg=None):
    """Exception Handler. Prints the executing function, failure and line no. Also writes error to opened log file and closes the log.\n
    param fnName: type(str). Executing function.
    param err: type(str). Details on the error.
    param lineNo: type(str). Line the error occured."""
    errorMsg = f"Error in executing Clawback Automation for the month:\n\
    Details: {errMsg}\n\
    Function name: {fnName}\n\
    Error: {err}\n\
    Line no: {lineNo}\n"
    print(errorMsg)
    write_to_logger(errorMsg)  
    processText = errorMsg

    layout = [
        [sg.Text(processText)],
        #[sg.Input(key="FILE"), sg.FileBrowse()],
        [sg.Button("Ok")]
    ]
    
    window = sg.Window("Clawback Process Error", layout, finalize=True, element_justification='c')
    
    while True:
        event, values = window.read()
        if event == "Ok":
            break    
        elif event == sg.WIN_CLOSED:
            break
    
    window.close()
    logFile.close()
    sys.exit()

def process_completed(process:str, exceptionCnt:int=None):
    """Generates a GUI with the process completed using the process key to map to the layout text message.\n
    param process: type(str). The running process key.
    param exceptionCnt: type(int). Count of noted exceptions. Only used for Clawback exceptions and ACH Exceptions"""
    exceptionCnt = str(exceptionCnt)
    processTextDict = {"clawback": "Clawback Process Completed",
                      "eligible": "Eligible Clawback File Generated, "+exceptionCnt+ " clawback exceptions noted",
                      "ach": "ACH File Received",
                       "noException": "No ACH Exceptions",
                       "achException": exceptionCnt+" ACH File Exceptions",
                       "timeout":'Timeout: No ACH File Received',
                       "gamers": "Unable to read latest Gamers Report. Using Gamers list on Master File"
                      }
    processText = processTextDict[process]
    #layoutText = "Clawback Process Completed"
    layout = [
        [sg.Text(processText)],
        #[sg.Input(key="FILE"), sg.FileBrowse()],
        [sg.Button("Ok")]
    ]
    
    window = sg.Window("Clawback Process", layout, finalize=True, element_justification='c')
    
    while True:
        event, values = window.read(timeout=10000)
        if event == "Ok":
            break    
        elif event == sg.WIN_CLOSED or event == '__TIMEOUT__':
            break
    
    write_to_logger(processText)
    window.close()



def tableDataText(table):    
    """Parses a html segment started with tag <table> followed 
    by multiple <tr> (table rows) and inner <td> (table data) tags. 
    It returns a list of rows with inner columns. 
    Accepts only one <th> (table header/data) in the first row.
    """
    def rowgetDataText(tr, coltag='td'): # td (data) or th (header)       
        return [td.get_text(strip=True) for td in tr.find_all(coltag)]  
    rows = []
    trs = table.find_all('tr')
    headerow = rowgetDataText(trs[0], 'th')
    if headerow: # if there is a header row include first
        rows.append(headerow)
        trs = trs[1:]
    for tr in trs: # for every table row
        rows.append(rowgetDataText(tr, 'td') ) # data row       
    return rows

def parse_exceptions_from_mail(folderPath:str, mailName:str, mailColumns:list):
    """Reads Exceptions emails and parses the email body to get the exception details.\n
    param folderPath: tye(str). PathName of the folder where the exception email is located.
    param mailName: type(str). Name of the email file (html format)."""

    inputFormat = '%Y_%m_%d_%H_%M_%S'
    outputFormat = '%Y-%m-%d %H:%M:%S'
    fileDate = datetime.datetime.strptime(mailName[-24:-5],inputFormat)    
    htmlFile = folderPath+mailName
    with open(htmlFile, encoding = "windows-1252") as file:
        soup = BeautifulSoup(file, "html.parser")
    htmltable = soup.find_all('table', { 'class' : "MsoNormalTable" })
    for idx in range(0, len(htmltable)):
        if idx == 0:
            list_table = tableDataText(htmltable[idx])
            exceptionMailDf = pd.DataFrame(list_table)
            exceptionMailDf.columns = exceptionMailDf.iloc[0]
            exceptionMailDf = exceptionMailDf.loc[1:]
            for col in [item for item in mailColumns if item not in list(exceptionMailDf.columns)]:
                exceptionMailDf[col] =np.nan
            exceptionMailDf = exceptionMailDf[mailColumns]
        else:
            list_table  = tableDataText(htmltable[idx])
            exceptionMailDf_2 = pd.DataFrame(list_table)
            exceptionMailDf_2.columns = exceptionMailDf_2.iloc[0]
            exceptionMailDf_2 = exceptionMailDf_2.loc[1:]
            for col in [item for item in mailColumns if item not in list(exceptionMailDf_2.columns)]:
                exceptionMailDf_2[col] =np.nan
            exceptionMailDf_2 = exceptionMailDf_2[mailColumns]
            exceptionMailDf = pd.concat([exceptionMailDf, exceptionMailDf_2], axis = 0)

    exceptionMailDf.reset_index(drop=True, inplace=True)
    exceptionMailDf.replace('â€¦','...', regex = True, inplace = True)
    exceptionMailDf["fileName"] = mailName
    exceptionMailDf["fileDate"] = fileDate
    
    return exceptionMailDf


def get_gamers_list(masterFilePath:str,gamersSheetName:str )->pd.DataFrame:
    """Returns the dataframe 'gamersDf' of known gamers. The function sources primarily from 
        the Gamers Report shared folder and secondarily from the Gamers sheet in Master File
        param masterFilePath: type(str). Pathname of the clawback masterfile selected.
        param gamersSheetName: type(str) Sheet name of the list of gamers in the masterfile"""
    
    try:
        gamersFileName = "Gamer's Report - Apr 2024.xlsm"
        gamersFilePath = ClawbackTestFolder+gamersFileName
        
        try:
            clawbackFilePath = ClawbackFiles
            folderPath = r"\\prdfile001\CIBC\RESERVE CLAWBACK\Gamers Report"            
            folderList = list(map(os.path.basename, glob.glob(clawbackFilePath+r"\Dealer Portfolio*.xlsm")))
            for fileName in folderList:
                cleanFileName = re.sub(r'_(.*?)\.', '.', fileName)
                my_file = Path(folderPath+"\\"+cleanFileName)
                if not my_file.is_file():
                    shutil.copy2(clawbackFilePath+fileName, folderPath)
                    os.replace(folderPath+fileName, folderPath+cleanFileName)
            folderList = list(map(os.path.basename, glob.glob(folderPath+r"\*.xlsm")))
            fileDict = {}
            for i in range(0,len(folderList)):
                try:
                    fileDict[folderList[i]] = parse('01 '+folderList[i].split('-')[-1].strip()[:-5])
                except Exception as e:
                    logMessage = "Unable to parse date from file "+ folderList[i] + ": "+ e +". \n Skipping file"
                    write_to_logger(logMessage)
                    continue
            earliestFile = sorted(fileDict.items(),key=lambda item: item[1], reverse = True)[0][0]
            gamersFileName = r"\\"+earliestFile
            gamersFilePath = r"\\prdfile001\CIBC\RESERVE CLAWBACK\Gamers Report"+gamersFileName
            gamersDf = pd.read_excel(gamersFilePath, sheet_name = "Gamer's Report" , header = 1)
            gamersDf = gamersDf[(gamersDf["Offenders"] == "Y") | (gamersDf["Worst Offenders"] == "Y") ].iloc[:,1:3]
            logMessage = "Loaded Gamers List from "+ gamersFilePath + " in Shared Folder"
                     
        except Exception as e:
            gamersDf = pd.DataFrame(columns=["Dealer ID", "Dealer Name"])
        
        if len(gamersDf)==0:
            process_completed('gamers')
            gamersDf = pd.read_excel(masterFilePath, sheet_name = gamersSheetName).iloc[:, 0:2]
            logMessage = "Loaded Gamers List from Master File"
        if len(gamersDf)==0:
            gamersDf = pd.read_excel(gamersFilePath, sheet_name = "Gamer's Report",header = 1)
            gamersDf = gamersDf[(gamersDf["Offenders"] == "Y") | (gamersDf["Worst Offenders"] == "Y") ].iloc[:,1:3]
            logMessage = "Loaded Gamers List from "+gamersFilePath+" in Backup ClawbackTest folder"
        
        gamersDf.columns = ["Dealer ID", "Dealer Name"]
        gamersDf.reset_index(drop = True, inplace = True)
        
        gamersDf = gamersDf.drop_duplicates()
        write_to_logger(logMessage)
        return gamersDf
    except Exception as f:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('get_gamers_list', f, exc_tb.tb_lineno)

def get_letter_version_list(masterFilePath:str, eftChequeSheetName:str, masterSheetName:str)->pd.DataFrame:
    """Generates a dataframe of dealers and their clawback letter version.
        Returns a dataframe 'eftChequeDf3'\n
        param masterFilePath: type(str). Pathname of the Clawback masterfile selected
        eftChequeSheetName: type(str) Sheet name of the letter version sheet in the Clawback masterfile
        masterSheetName: type(str). Sheet name of the recods of clawbacks in the clawback masterfile."""

    try:
        ##Read Dealers' Letter Version from eftChequeSheetName in masterfile and drop duplicates
        eftChequeDf = pd.read_excel(masterFilePath, sheet_name = eftChequeSheetName).iloc[:,:6]
        eftChequeDf = eftChequeDf[["ORIGINATOR CODE TXT","ORIGINATOR NAME TXT", "NEW EFT RECVD?" ]]
        eftChequeDf.rename(columns ={'ORIGINATOR CODE TXT':'Dealer ID','NEW EFT RECVD?': 'Letter Version', "ORIGINATOR NAME TXT":'Dealer'}, inplace = True)
        eftChequeDf=eftChequeDf.drop_duplicates()

        ##Read Dealers' Letter Version from historic masterSheetName in masterfile. Rank and drop duplicates 
        eftChequeDf2 = pd.read_excel(masterFilePath, sheet_name = masterSheetName)[["Dealer ID", "Dealer", "Letter Version"]]
        eftChequeDf2['RN'] = eftChequeDf2.reset_index().sort_values(['index'], ascending=[False]) \
                            .groupby(['Dealer ID']) \
                            .cumcount() + 1
        eftChequeDf2 = eftChequeDf2[eftChequeDf2['RN'] == 1]
        eftChequeDf2 = eftChequeDf2[["Dealer ID", "Dealer", "Letter Version"]]

        ##Combine both Letter Version dataframes. Rank and drop duplicates
        eftChequeDf = pd.concat([eftChequeDf2,eftChequeDf], axis = 0).drop_duplicates().reset_index(drop = True)
        eftChequeDf['RN'] = eftChequeDf.reset_index().sort_values(['index'], ascending=[False]) \
                            .groupby(['Dealer ID']) \
                            .cumcount() + 1
        eftChequeDf3 = eftChequeDf[eftChequeDf['RN'] == 1]
        eftChequeDf3.rename(columns = {"Dealer":"Letter Dealer"}, inplace = True)
        return eftChequeDf3
    except Exception as f:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('get_letter_version_list', f, exc_tb.tb_lineno)


def check_gamers(dealerId:str, daysOnBook:int, gamersDf:pd.DataFrame):
    """Checks if a given dealer is eligible for exception. Dealers not eligible include gamers to the clawback process using and if the days on book is in violation.\n
    param dealerId: type(str). Dealer ID.
    param daysOnBook: type int. Deal's days on book.
    param gamersDf: type(pd.DataFrame). Dataframe of known gamers."""
    if dealerId in list(gamersDf['Dealer ID']):
        out = "Gamer list"
    elif daysOnBook < 40:
        out = "Less than 40 days"
    elif daysOnBook > 170:
        out = "More than 170 days"
    else:
        out = 'Eligible for exception'
    return out

def letter_eft(dealerId:str, letterEft, activationDate:datetime.datetime):
    """Checks if a dealer's clawback is to be carried out by EFT or Cheque based on preset rules.\n
    param dealerId: type(str). Dealer ID.
    param letterEft: type(str). Dealer Letter Version. Only applicable for Ontario dealers.
    param activationDate: type(datetime.datetime). Dealer activation date.\n
    Rules
    1. If the dealer is based in Ontario, function references the letterEFT dataframe.\n
    2. If the dealer is based in Ontario and is not in the letterEFT dataframe, 
    function references the dealer activation date - after 2017-01-01 is EFT, else TBD. \n
    3. If the dealer is based in Quebec, then QC EFT.\n
    4. If the dealer activation date is after 2017-01-01 then EFT.\n
    Any other case is treated on an exceptional basis\n    
    """
    if type(letterEft) == str:
        out = letterEft
    elif dealerId[0:2] == 'ON' and type(letterEft)!=str:
        if activationDate >= parse('01 Jan 2017'):
            out = "EFT"
        else:
            out = None
    elif dealerId[0:2] == 'QC':
        out =  "QC EFT"
    elif activationDate >= parse('01 Jan 2017') :
        out = "EFT"
    elif dealerId == 'AB0186':
        out = "EFT"
    else:
        out = None
    return out


def get_masterfile_file_and_date():
    """Creates GUI for selecting Clawback masterfile and the clawback month."""
    #window = sg.Listbox(dateList, size=(20,4), enable_events=False, key='_LIST_')
    file = None
    date = None
        
    dateList = pd.date_range(start = parse('01-Jan-2022'), end = datetime.datetime.today()+ pd.offsets.MonthBegin(0), freq = 'MS')
    dateListStr = []
    for x in dateList[:-1]:
        y = x.date().strftime('%b %Y')
        dateListStr.append(y)
    dateListStr.reverse()
    
    layout = [
        [sg.Text("Select File and Month")],
        [sg.Input(key="FILE"), sg.FileBrowse()],
        [sg.Listbox(dateListStr, size=(20,10), key='LISTBOX')],
        [sg.Button("Read")]
    ]
    
    window = sg.Window('Title', layout, finalize=True)
    #listbox = window['LISTBOX']
    
    while True:
        event, values = window.read() ###event, values = window.read(timeout=50000)
        if event == "Read":
            date = values["LISTBOX"]
            file = values["FILE"]
            #print(date)
            if os.path.exists(file) and len(date) !=0:
                logMessage = "Master File and Date selected"
                write_to_logger(logMessage)
                break    
            elif event == sg.WIN_CLOSED:
                break
        
        if event == sg.WIN_CLOSED: #### or event == '__TIMEOUT__':
            break
    
    window.close()
    
    return date, file
    
    

def clawback_for_ach():
    """
    The function `clawback_for_ach` retrieves a master file and date, logs the selections, and returns
    the clawback date and master file title.
    :return: The function `clawback_for_ach` is returning the `clawbackDate` and `masterFile` variables
    if certain conditions are met. If the `file` is not None and the `date` list is not empty, the
    function will return the `clawbackDate` and `masterFile` values. Otherwise, it will return `None,
    None`.
    """
    date, file = get_masterfile_file_and_date()

    if type(file) ==type(None) or len(date) ==0 or len(file) ==0:
        return None, None
    else:
        try:
            masterFile = file.title()
            logMessage = "Master File selected: "+ masterFile
            write_to_logger(logMessage)

            clawbackDate = date[0]
            logMessage = "Clawback Date selected: " + clawbackDate
            write_to_logger(logMessage)

            return clawbackDate, masterFile
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            raise_fn_exception('clawback_for_ach', e, exc_tb.tb_lineno)


def get_letter_version_exceptions(letterVersionExceptions:pd.DataFrame, eftChequeDf:pd.DataFrame, clawbackColumns:list, dt:datetime.datetime):
    
    fileDate = dt.strftime('%Y_%m_%d')
    try:
        letterVersionExceptions = letterVersionExceptions.merge(eftChequeDf[eftChequeDf['Dealer ID'].str[0:2] =='ON'], how = 'left', on = ["Dealer ID"])                
        for row in list(letterVersionExceptions.index):
            letterVersionExceptions.loc[row, 'Letter Version.1'] = letter_eft(letterVersionExceptions.loc[row,'Dealer ID'], letterVersionExceptions.loc[row, 'Letter Version_y'], parse("01 Jan 2016"))
        letterVersionExceptions['Letter Version'] = letterVersionExceptions['Letter Version.1']
        logMessage = "Letter Version exceptions found. See letter version exceptions file for details"
        letterVersionExceptions[clawbackColumns].to_csv(ClawbackOutFolder+"Letter_version_exceptions_"+fileDate+".csv")
        write_to_logger(logMessage)
        return letterVersionExceptions[clawbackColumns]
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('get_letter_version_exceptions', e, exc_tb.tb_lineno)


def final_eligible_clawback(date:str, masterFile:str):
    """Prepares the reports for the final clawback process, including final version of Clawback Masterfile for the month (with exceptions) and Eligible Clawbacks for ACH file generator.\n
    param date: type(str). Selected Month-year of Final Clawback/ACH File Generation process.
    param masterFile: type(str). File Path of the Selected Masterfile of the Final Clawback/ACH File Generation process"""
    
    
    dt = parse("01 "+date)
    fileDate = dt.strftime('%Y_%m_%d')
    fileDateString = dt.strftime('%B, %Y')

    folderPath = ClawbackOutFolder
    eligibleClawbackCsv = folderPath+"ELIGIBLE CLAWBACKS 2.xlsx"
    monthExceptionsCsv = folderPath+"Clawback Exceptions "+fileDate+".xlsx"
    newMonthFilePath = folderPath+ "CIBC RESERVE CLAWBACK MASTER FINAL "+fileDateString+".xlsx"
    trendCsv = folderPath+"Clawback_Summary_Trend_"+fileDate+".xlsx"

    masterSheetName = "ELIGIBLE CLAWBCKS"
    

    clawbackColumns =   ['CLASS #',	 'CIBC Client ID',	 
                        'Vehicle VIN#','Dealer','Client Name',
                        'Disbersement Date','Loan Amount','Rate',	 
                        'Payout Date','Payout Month',	 
                        'Days on book','Reserve Amount',	 
                        'Dealer ID','CIBC Sales Rep',	 
                        'PortalID','Letter Version',	 
                        'Letter Version.1','Ineligible Dealers List',	 
                        'Dealer Operating Name',
                        'Exception %','Exception Amount',
                        'Clawback Amount','Amount Received',
                        'Final Classification','Exception Reason',
                        'Notes','Email','Commercial']

    newMasterColumnMap = {'Customer Name':'Client Name', 
                       'Reserve Amount ':'Reserve Amount',
                        'Disbersement Date ':'Disbersement Date',
                         'Sales Rep':'CIBC Sales Rep',
                         'Portal Number':'PortalID' 
                         }
    try:
        
        
        masterFilePath = masterFile
        masterDf = pd.read_excel(masterFilePath, sheet_name = masterSheetName).iloc[:,1:]       
        masterMaxRow = len(masterDf)

        newMasterDf=masterDf[masterDf["Payout Month"]==dt]
        newMasterDf.rename(newMasterColumnMap, axis = 1, inplace = True) 

        masterMaxRow = masterMaxRow - len(newMasterDf)+1
        
        monthExceptions = exceptions_master(masterDf, masterSheetName, monthExceptionsCsv)
        exceptionsCnt = len(monthExceptions[monthExceptions["Payout Month"]==dt.date()])
        
        left_a = newMasterDf.set_index(['PortalID', 'Payout Month'])
        right_a = monthExceptions.rename({'Portal Number':'PortalID'}, axis = 1).iloc[:,:-3].set_index(['PortalID', 'Payout Month'])
        res = left_a.reindex(columns=left_a.columns.union(right_a.columns))
        res.update(right_a)
        res.reset_index(inplace=True)

        ##clawbackColumns = clawbackColumns+exceptionColumns
        clawback2 = res[clawbackColumns]
        clawback2['Exception Amount'] = clawback2['Exception %']*clawback2['Reserve Amount']
        clawback2['Clawback Amount'] = clawback2['Reserve Amount']-clawback2['Exception Amount']
        clawback2['Amount Received']=clawback2['Clawback Amount']
        clawback2['Payout Month'] = clawback2['Payout Month'].apply(lambda x:pd.to_datetime(int(x/1e6), utc=True, unit='ms').date() if type(x) == int else x)
        clawback2["Payout Date"] = pd.to_datetime(clawback2["Payout Date"])
        clawback2["Payout Month"] = pd.to_datetime(clawback2["Payout Month"])
        clawback2['CIBC Client ID'] = clawback2['CIBC Client ID'].astype('int')

        letterVersionExceptions = clawback2[clawback2["Letter Version"].isnull()]
        if len(letterVersionExceptions) != 0:

            eftChequeSheetName = "DEALER LIST - EFT VS CHEQUE"
            eftChequeDf = get_letter_version_list(masterFilePath, eftChequeSheetName, masterSheetName)   
            letterVersionExceptions = get_letter_version_exceptions(letterVersionExceptions, eftChequeDf, clawbackColumns, dt)
            left_a = clawback2.set_index(['PortalID', 'Payout Month'])
            right_a = letterVersionExceptions.set_index(['PortalID', 'Payout Month'])
            res = left_a.reindex(columns=left_a.columns.union(right_a.columns))
            res.update(right_a)
            res.reset_index(inplace=True)
            clawback2 = res[clawbackColumns]
        
        outSheetName = "ACH - "+(dt + pd.tseries.offsets.MonthEnd(2)).strftime('%B %Y')
        

        eligibleClawback = clawback2[((clawback2["Letter Version"] == "EFT") | (clawback2["Letter Version"] == "QC EFT")) & ((clawback2["Exception %"] != 1) | ((clawback2["Exception %"]).isnull()))]
        eligibleClawback["Message"] = 'CIBC Auto Finance'
        eligibleClawback['Clawback Pull Date'] = get_last_working_date((dt + pd.tseries.offsets.MonthEnd(2))).date()
        eligibleClawback['Clawback Amount'] = round(eligibleClawback['Clawback Amount'],2)
        eligibleClawback = eligibleClawback[['PortalID','CIBC Client ID',  'Client Name', "Dealer", "Dealer ID", "Message", "Clawback Amount",  "Clawback Pull Date"]]
            
        eligibleClawback.columns =  ["Portal ID","CMSI App ID","Client Name","Dealer Name","Dealer ID","Message","Clawback Amount","Clawback Pull Date" ]
        
        eligibleClawback.to_excel(eligibleClawbackCsv, index = False, sheet_name = outSheetName)
        
        
        newMasterDf = write_to_master(masterFilePath, newMonthFilePath, masterSheetName, clawback2, masterMaxRow)
        trendDf = generate_trend_report(newMasterDf,trendCsv)

        process_completed("eligible", exceptionsCnt)
        logMessage = "Final eligible clawback files for "+date+" created successfully."
        write_to_logger(logMessage)

        return eligibleClawback, monthExceptions, clawback2
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('final_eligible_clawback', e, exc_tb.tb_lineno)

def get_ach_exceptions(eligibleClawback:pd.DataFrame):
    """Listens for ACH File generated and sent to user email, then compares the ACH file received with the Eligible Clawback file sent for missing dealers.
    param eligibleClawback: type(pd.DataFrame). Deals eligible for clawback via EFT."""

    folderPath = ClawbackFiles
    outFolderPath = ClawbackOutFolder
    archiveFolder = ClawbackFiles+"Archive/"
    
    
    searchword = "INO_FALCON_CLAWBACK_ACH"
    try:
        
        cnt = 1
        while True:
            logMessage = "Listening for ACH file: "+str(cnt)+" min"
            write_to_logger(logMessage)
            if cnt > 10:
                message= 'timeout'
                return message
            else:
                folderList = list(map(os.path.basename, glob.glob(folderPath+"*.txt")))
                searchFiles = []
                for item in folderList:
                    if item.startswith(searchword):
                        searchFiles.append(item)
                
                
                if len(searchFiles)>0:
                    process_completed('ach')
                    break
                else:
                    cnt = cnt+1
                    time.sleep(60)
                
        for report in searchFiles:
            inputFormat = '%Y-%m-%d %H_%M_%S'
            outputFormat = '%Y-%m-%d %H_%M_%S'

        
            fileName = report
            fileDate = datetime.datetime.strptime(fileName[-23:-4],inputFormat)
            fileDateString = fileDate.strftime(outputFormat) 

            achExceptionsCsv = outFolderPath+"ACH Exceptions_"+fileDateString+".xlsx"   
            
            with open(folderPath+fileName) as f:
                df = pd.DataFrame(f)
            
            df2 = []
            for i in range(0,len(df)):
                x = [x.strip() for x in df.loc[i][0].split(r"CIBC AUTO FINANCE ")]
                df2.append(x)
            
            df3 = pd.DataFrame(df2)
            df3["CMSI App ID"] = df3[1].apply(lambda x: (x.split(' ')[0][10:]).strip() if x != None else x)
            df3 = df3[~df3["CMSI App ID"].isnull()]
            df3["CMSI App ID"] = df3["CMSI App ID"].astype('int64')
            eligibleClawback = eligibleClawback.merge(right = df3, how = 'left', on ='CMSI App ID')
            achExceptions = eligibleClawback[eligibleClawback[0].isnull()]
            achExceptions.to_excel(achExceptionsCsv)

            shutil.copy2(folderPath+fileName, archiveFolder+fileName)
            os.remove(folderPath+fileName)
            return achExceptions
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('get_ach_exceptions', e, exc_tb.tb_lineno)
    


def file_select_window(fileSelect:str):
    """Creates a GUI for user to select a file. Text prompt depends on the 'fileSelect' string variable, which maps to a dictionary. Returns the file selected\n
    param fileSelect: type(str). Key for the layout text to be displayed in the GUI window"""

    file = None
    
    promtDict = {"clawback file": "Select Clawback Source File for the month",
                "master file": "Select previous month Master File"}

    layoutText = promtDict[fileSelect]
    layout = [
        [sg.Text(layoutText)],
        [sg.Input(key="FILE"), sg.FileBrowse()],
        [sg.Button("Read")]
    ]
    
    window = sg.Window("Read from Excel file", layout, finalize=True)
    
    while True:
        event, values = window.read()
        if event == "Read":
            file = values["FILE"]
            if os.path.exists(file):
                break    
            elif event == sg.WIN_CLOSED:
                break
        if event == sg.WIN_CLOSED:
            break

    window.close()
    return file

def initiate_clawback_files(filePath:str, masterFilePath:str):
    """Loads all files required to generate the monthly clawbacks.\n
    param filePath: type(str). The filepath of the Clawback source file for the month.
    param masterFilePath: type(str). The filepath of the Clawback Master file for the previous month.\n"""
    
        
    masterSheetName = "ELIGIBLE CLAWBCKS"
    gamersSheetName = "Gamers"
    eftChequeSheetName = "DEALER LIST - EFT VS CHEQUE"
    
    sfMasterFileName = "Salesforce Email Master File.xlsx"
    sfMasterFilePath = ClawbackTestFolder+sfMasterFileName
    sfMasterSheetName = "Masterfile"
    

    columnMap = {'Class #': 'CLASS #',
    'Client APP ID':'CIBC Client ID',
    'VIN #':'Vehicle VIN#',
    'Dealer':'Dealer',
    'BRNDLR':'Dealer ID',
    'Disburse Date':'Disbersement Date',
    'Loan Amount':'Loan Amount',
    'Rate':'Rate',
    'Status Update Date':'Payout Date',
    'Days Paid':'Days on book',
    'RESERVE_AMT':'Reserve Amount'} 

    innovatecDealerscolumnMap = {
                            "Code": "Dealer ID",
                            "DealerID": "Innovatec Dealer Code",
                            "LegalName":"Innovatec LegalName",
                            "Name": "Innovatec DealerName",
                            "DealerStatusText": "Innovatec DealerStatus",
                            "InactiveReason":"Innovatec InactiveReason",
                            "DealerTypeText": "Innovatec DealerType",
                            "ManufacturerText":"Innovatec Manufacturer",
                            "SignupDate":"Innovatec SignupDate",
                            "PhoneNumber":"Innovatec PhoneNumber",
                            "AddressStreet":"Innovatec Street Address",
                            "City":"Innovatec City",
                            "Province":"Innovatec Province",
                            "PostalCode":"Innovatec Postal Code",
                            "Country":"Innovatec Country",
                            "EFTEmail":"Innovatec EFTEmail",
                            "Email1":"Innovatec Email1",
                            "DBAName":"Innovatec DBAName",
                            "EFTFaxNumber":"Innovatec EFTFaxNumber",
                            "BankNumber":"Innovatec BankNumber",
                            "BankBranch":"Innovatec BankBranch",
                            "BankAccountNumber":"Innovatec BankAccountNumber"}

    newDealerCode =  {'BC5282' : 'BC0416',
                    'BC5283' : 'BC0419',
                    'BC5284' : 'BC0420',
                    'BC5285' : 'BC0421',
                    'BC5286' : 'BC0422',
                    'BC5287' : 'BC0423',
                    'BC5288' : 'BC0417',
                    'BC5289' : 'BC0424',
                    'BC5290' : 'BC0425',
                    'BC5292' : 'BC0426',
                    'BC5293' : 'BC0427',
                    'BC5294' : 'BC0428',
                    'BC5295' : 'BC0429',
                    'BC5296' : 'BC0430',
                    'BC5297' : 'BC0418',
                    'BC5798' : 'BC0432',
                    'BC5290' : 'BC0425',
                    'BC5402' : 'BC0431',
                    'BC5798' : 'BC0432'
                    } 
    
    sfMapper = {'ORIGINATOR CODE TXT':'Dealer ID',
                'DLR ADDRESS1 TXT':'Street Address',
                'DLR CITY TXT':'City',
                 'DLR STATE ID':'Province',
                 'DLR ZIPCODE TXT':'Postal Code',
                 'EMAIL ADDRESS TXT':'Clawback Email'}
    

    cibcCubequery = text('with clientID_data  as ( \
                                Select  [ClientAppId], [PortalID], \
                                [Client Name], [Client Email], [applicationid], [dealerID],\
                                [Dealer Name], \
                                row_number() over (partition by [ClientAppId] order by a.[InitiationDate] desc ) as rnk\
                                from [CIBC].[dbo].[cibc_cub_dataset_F] a)\
                                select cast([ClientAppId] as int) as ClientAppId,\
                                [PortalID], [Client Name], [Client Email], \
                                [applicationid], [dealerID], [Dealer Name]\
                                from clientID_data\
                                where rnk = 1')
    
    innovatecDealersQuery  = text("with cibcDealers as (SELECT [Code]\
      ,[DealerID],[LegalName],[Name],[DealerStatusText],[InactiveReason]\
      ,[DealerTypeText],[ManufacturerText],[SignupDate],[PhoneNumber]\
      ,ltrim(rtrim(concat([AddressNumber],' ', [AddressStreet]))) as AddressStreet\
      ,[City],[Province],[PostalCode],[Country],[EFTEmail],[Email1]\
      ,[DBAName],[EFTFaxNumber],[BankNumber],[BankBranch]\
      ,[BankAccountNumber]\
      , row_number() over (partition by code order by [SignupDate] desc) as rnk\
  FROM [InovatecCIBC].[dbo].[Dealers])\
  select * from cibcDealers\
  where rnk = 1")
    
    dealerBodyQuery = text('with dealerBody as (select [Account Owner],[Account Name] ,[CIBC Activation Date] ,\
                            [CIBC Reporting ID] ,[Parent Account] ,[Dealer Group (TBC)] ,\
                            [CIBC Sales Rep] ,[Account ID] ,[Reporting Channel] ,[Franchise] ,\
                            [Account Owner Alias] ,[Imperial ID] ,[CIBC Quad] ,[Fax Number] ,\
                            [Dealer Type] ,[Reserve Clawback Email] ,[Dealer State/Province] ,\
                            [CIBC Dealer Status] ,[Portal Preference], \
                            [Dealer Street], [Dealer City],[Dealer Zip/Postal Code] ,[Legal Name], \
                            row_number() over (partition by [CIBC Reporting ID] order by [CIBC Activation Date] desc, len([Dealer Zip/Postal Code]) desc) rnk \
                            FROM [Test].[dbo].[CIBC_Dealer_Body]) \
                            select * from dealerBody a \
                            where rnk = 1'
                              )
        

    try:
        ##File Date
        dt = parse('1 '+' '.join(pd.read_excel(filePath, header = 0).iloc[0:1,0].str.split(' ')[0][::5]))        

        ##SQL Engine and Query
        engine = connectionEngine()
        testEngine = connectionTestEngine()        

        ##Infiles
        df = pd.read_excel(filePath, header=5).iloc[:,1:]
        gamersDf = get_gamers_list(masterFilePath, gamersSheetName) 
        
        dealerBody = pd.read_sql_query(dealerBodyQuery, con = testEngine)
        masterDf = pd.read_excel(masterFilePath, sheet_name = masterSheetName).iloc[:,1:]
        eftChequeDf = get_letter_version_list(masterFilePath, eftChequeSheetName, masterSheetName)   
        cibcDf = pd.read_sql_query(cibcCubequery, con = engine)
        innovatecDealers = pd.read_sql_query(innovatecDealersQuery, con = engine)
        
        try:
            sfMasterFileName = r"\Salesforce Email Master File.xlsx"
            sfMasterFilePath = r"\\prdfile001\CIBC\RESERVE CLAWBACK"+sfMasterFileName
            sfMasterFile = pd.read_excel(sfMasterFilePath, sheet_name = sfMasterSheetName)            
        except:
            sfMasterFile = pd.read_excel(sfMasterFilePath, sheet_name = sfMasterSheetName)       
        

        

        ##Clean df dataframe
        df['Payout Month'] = dt
        df.rename(columns=columnMap, inplace = True) 
        df['Dealer ID'] = df['Dealer ID'].replace(newDealerCode)
        for col in ['Loan Amount','Rate','Reserve Amount']:
            df[col] = df[col].astype('float64')
        df['Days on book'] = df['Days on book'].astype('int64')
        df.set_index('CIBC Client ID', inplace = True)
        
        ##Clean CIBC DealerBody Dataframe
        dealerBody.rename(columns ={'CIBC Reporting ID':'Dealer ID', 'Account Name': 'Dealer Operating Name'}, inplace = True)

        ##Clean CIBC Dealers Dataframe
        innovatecDealers.rename(columns =innovatecDealerscolumnMap, inplace = True)

        ##Clean CIBC Cube Dataframe
        cibcDf.set_index("ClientAppId", inplace = True)

        ##Clean SalesForce Email Master File
        sfMasterFile.rename(columns=sfMapper, inplace = True)
        sfMasterFile['Dealer Contact'] = sfMasterFile['Province'].apply(lambda x: "A qui de droip:" if x=='QC' else "To whom it may concern")
       
        logMessage = "Initiating Clawback Files completed"
        write_to_logger(logMessage)
        return df, gamersDf, dealerBody, masterDf, eftChequeDf, cibcDf, sfMasterFile, innovatecDealers, dt
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('initial_eligible_clawback', e, exc_tb.tb_lineno)

def exceptions_master(masterDf:pd.DataFrame, masterFileName:str, exceptionsCsv:str):
    """Returns a dataframe 'exceptionsDf' of clawback exceptions. Function combines the parsed exceptions from the exception emails with the recorded exceptions in the clawback masterfile selected. Exports exceptions to ClawbackOutFolder as 'exceptionsCsv'\n
    param masterDf: type(pd.DataFrame). Dataframe of the selected clawback masterfile.
    param masterFileName: type(str). Name of the clawback masterfile.
    param exceptionsCsv: type(str). Pathname of the exceptions output file (.xlsx format)\n"""
    dateFormat = '%Y_%m_%d_%H_%M_%S'
    exceptioncolumns = ['Vehicle VIN#', 'Payout Month', "Portal Number", "CIBC Client ID", 'Dealer ID', "Dealer",
                        'Exception %',	'Exception Amount',	
                        'Clawback Amount',
                        'Amount Received',
                        'Final Classification',	
                        'Exception Reason',	
                        'Notes']
    
    try:
        exceptionsDf = masterDf[exceptioncolumns]
        exceptionsDf = exceptionsDf[~(exceptionsDf['Exception Reason']).isnull()]
        exceptionsDf["fileName"] = masterFileName
        exceptionsDf["fileDate"] = datetime.datetime.now()
        replacers = ['\$', ',']

        folderPath = ClawbackFiles
        folderList = list(map(os.path.basename, glob.glob(folderPath+"*.html")))
        for fileName in folderList:
            mailName = fileName
            mailFilePath=folderPath+mailName
            exceptionMailDf = parse_exceptions_from_mail(folderPath, mailName, exceptioncolumns)
            for col in exceptioncolumns:
                if col not in exceptionMailDf.columns:
                    exceptionMailDf[col] = np.nan
            exceptionMailDf = exceptionMailDf[exceptioncolumns].replace('',np.nan)
            exceptionMailDf= exceptionMailDf[~exceptionMailDf['Portal Number'].isnull()]
            exceptionMailDf['Exception %'].replace({'%':''}, inplace = True, regex = True)
            exceptionMailDf['Exception %'] = (exceptionMailDf['Exception %'].astype('float'))/100
            exceptionMailDf['Exception Amount'] = (exceptionMailDf['Exception Amount'].astype('str').apply(lambda x:re.sub('|'.join(replacers), '', x)).astype('float'))
            exceptionMailDf['Clawback Amount'] = (exceptionMailDf['Clawback Amount'].astype('str').apply(lambda x:re.sub('|'.join(replacers), '', x)).astype('float'))  ###apply(lambda x:str.replace(x,'$','')).astype('float'))
            exceptionMailDf['Amount Received'] = (exceptionMailDf['Amount Received'].astype('str').apply(lambda x:re.sub('|'.join(replacers), '', x)).astype('float'))
            exceptionMailDf["Exception Reason"] = exceptionMailDf["Exception Reason"].replace("relationship","Waived for Relationship")
            if (exceptionMailDf['Payout Month']).dtypes == 'O':
                try:
                    exceptionMailDf['Payout Month'] = exceptionMailDf['Payout Month'].apply(lambda x: parse('1 '+ x))                    
                except:
                    exceptionMailDf['Payout Month'] = exceptionMailDf['Payout Month'].apply(lambda x: parse(x))
            exceptionMailDf["fileName"] = fileName
            exceptionMailDf["fileDate"] = datetime.datetime.strptime(fileName[-24:-5],dateFormat)
            #exceptionMailDf['Payout Month']=exceptionMailDf['Payout Month'].dt.date
            exceptionsDf = pd.concat([exceptionsDf,exceptionMailDf], axis = 0)
        
        exceptionsDf.reset_index(drop=True, inplace = True)
        exceptionsDf['RN'] = exceptionsDf.sort_values(['Payout Month', 'fileDate'], ascending=[False, False]) \
                    .groupby(['Portal Number', 'Payout Month']) \
                    .cumcount() + 1
        exceptionsDf = exceptionsDf[exceptionsDf['RN']==1].sort_values(['Payout Month', 'fileDate', 'Vehicle VIN#'], ascending=[True, True, True]).reset_index(drop = True)
        exceptionsDf['Payout Month']=exceptionsDf['Payout Month'].dt.date
        exceptionsDf["Notes"] = exceptionsDf.apply(lambda x: x["Exception Reason"] if not(str.lower(x["Exception Reason"]) in ("waived for relationship", "branch solicitation", "waived for insurance write off", "other (rebooks, 1-off situations)")) and pd.isnull(x["Notes"]) else x["Notes"], axis = 1)
        exceptionsDf['Final Classification'].fillna("Exception", inplace = True)
        exceptionsDf.to_excel(exceptionsCsv) ###, encoding = "windows-1252") 
        logMessage = "Exceptions file generated"
        write_to_logger(logMessage)
        return exceptionsDf
    except Exception as f:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('exceptions_master', f, exc_tb.tb_lineno)

def eligible_clawback(df:pd.DataFrame,cibcDf:pd.DataFrame,dealerBody:pd.DataFrame,eftChequeDf:pd.DataFrame,gamersDf:pd.DataFrame,exceptionsDf:pd.DataFrame, dt:datetime.datetime, clawbackCsv:str):
    """Returns a dataframe 'clawback2' of clawbacks using the source file for the month. Saves clawback2 dataframe as an excel file to the ClawbackOutFolder as clawbackCsv. \n
    param df: type(pd.dataFrame). source file for the month.
    param cibcDf: type(pd.DataFrame). CIBC cube dataset.
    param dealerBody: type(pd.DataFrame). CIBC dealerbody report.
    param eftChequeDf: type(pd.DataFrame). EFT/Checque worksheet
    param gamersDf: type(pd.DataFrame). Gamers dataframe 
    param exceptionsDf: Exceptions dataframe.
    param dt: type(datetime.datetime). Clawback month in y-M-d format.
    param clawbackCsv: type(str). Pathname of the clawback export file"""
    
           
    clawbackColumns = ['CLASS #',	 'CIBC Client ID',	 
                      'Vehicle VIN#','Dealer','Dealer Name','Client Name',
                      'Disbersement Date','Loan Amount','Rate',	 
                      'Payout Date','Payout Month',	 
                      'Days on book','Reserve Amount',	 
                      'Dealer ID','CIBC Sales Rep',	 
                      'PortalID','Letter Version',                      	 
                      'Dealer Operating Name','CIBC Activation Date',
                       'Reserve Clawback Email',
                        'Dealer State/Province',
                       'CIBC Dealer Status' ,'Portal Preference', 
                       'Dealer Street', 'Dealer City',
                       'Dealer Zip/Postal Code' ,'Legal Name']

    exceptionColumns = ['Letter Version.1',
                        'Ineligible Dealers List','Exception %',
                        'Exception Amount',
                      'Clawback Amount','Amount Received',
                      'Final Classification','Exception Reason',
                      'Notes','Email','Commercial']
    
    try:
        clawback = df.merge(cibcDf, how = 'left',  left_index=True, right_index=True )
        clawback.reset_index(drop=False, inplace = True)
        clawback = clawback.merge(dealerBody, how = 'left', on = ["Dealer ID"])
        clawback = clawback.merge(eftChequeDf[eftChequeDf['Dealer ID'].str[0:2] =='ON'], how = 'left', on = ["Dealer ID"])
        clawback['CIBC Activation Date'] = clawback['CIBC Activation Date'].astype('<M8[ns]')

        clawback2 = clawback[clawbackColumns]  
        clawback2[exceptionColumns] = np.nan

        for row in range(0,len(clawback2)):
            clawback2.loc[row, 'Ineligible Dealers List']=check_gamers(clawback2.loc[row,'Dealer ID'], clawback2.loc[row, 'Days on book'], gamersDf)
        
        for row in range(0,len(clawback2)):
            clawback2.loc[row, 'Letter Version.1']=letter_eft(clawback2.loc[row,'Dealer ID'], clawback2.loc[row, 'Letter Version'], clawback2.loc[row,'CIBC Activation Date'])

        clawback2['Letter Version'] = clawback2['Letter Version.1']
        clawback2['Exception %'] = 0

        #print("Milestone")

        left_a = clawback2.set_index('Vehicle VIN#')
        right_a = exceptionsDf.iloc[:,:-3].set_index('Vehicle VIN#')
        res = left_a.reindex(columns=left_a.columns.union(right_a.columns))
        res.update(right_a)
        res.reset_index(inplace=True)

        clawbackColumns = clawbackColumns+exceptionColumns
        clawback2 = res[clawbackColumns]
        clawback2['Exception Amount'] = clawback2['Exception %']*clawback2['Reserve Amount']
        clawback2['Clawback Amount'] = clawback2['Reserve Amount']-clawback2['Exception Amount']
        clawback2['Amount Received']=clawback2['Clawback Amount']
        clawback2['Message'] = 'CIBC Auto Finance'
        clawback2['Clawback Pull Date'] = get_last_working_date((dt + pd.tseries.offsets.MonthEnd(2))).date()

        letterVersionExceptions = clawback2[clawback2["Letter Version"].isnull()]
        if len(letterVersionExceptions) != 0:
            letterVersionExceptions = get_letter_version_exceptions(letterVersionExceptions, eftChequeDf, clawbackColumns, dt)

        outSheetName = "ACH - "+(dt + pd.tseries.offsets.MonthEnd(2)).strftime('%B %y')
        
        clawback2.to_excel(clawbackCsv, sheet_name = outSheetName, index = False) ####encoding = "windows-1252"

        logMessage = "Eligible Clawbacks File generated successfully."
        write_to_logger(logMessage)
        return clawback2
    except Exception as g:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('eligible_clawback', g, exc_tb.tb_lineno)

def write_to_master(masterFilePath:str, newMonthFilePath:str, masterSheetName:str, clawback2:pd.DataFrame, masterMaxRow:int):
    """Creates a copy of the Masterfile 'masterFilePath' as 'newMonthFilePath' in ClawbackOutFolder and inserts into it data from the 'clawback2' dataframe.
    Insert is made to sheet 'masterSheetName' and at row number given by 'masterMaxRow'. Returns the updated newMonthFilePath as 'newMasterDf'\n
    param newMonthFilePath: type(str). New Clawback Masterfile for the month.
    param masterFilePath: type(str). Selected Clawback masterfile workbook. For Generate Clawback File process, it is the master file for previous month. For Generate ACH File process, it is the master file for the current month.
    param masterSheetName: type(str). Name of the MasterFile sheet in the masterfile workbook.
    param clawback2: type(pd.DataFrame). Dataframe of the clawbacks to be inserted to the newMonthFilePath masterfile.
    param masterMaxRow: type(int). Sheet row on which the clawback2 dataframe is inserted. For Generate Clawback File process, it is the first empty row on the sheet. For Generate ACH File process, it is the first row for the selected payout month and year. """
    try:
        #fileDate = datetime.datetime.strptime(fileDate,'%Y_%m_%d').strftime('%B, %Y')
        ##masterFilePath = folderPath+masterFile
        
        if os.path.isfile(newMonthFilePath):
            os.remove(newMonthFilePath)
        shutil.copy2(masterFilePath, newMonthFilePath)
        masterColumns = ['CLASS #',	 'CIBC Client ID',	 
                        'Vehicle VIN#','Dealer','Client Name',
                        'Disbersement Date','Loan Amount','Rate',	 
                        'Payout Date','Payout Month',	 
                        'Days on book','Reserve Amount',	 
                        'Dealer ID','CIBC Sales Rep',	 
                        'PortalID','Letter Version',	 
                        'Letter Version.1','Ineligible Dealers List',	 
                        'Dealer Operating Name',
                        'Exception %','Exception Amount',
                        'Clawback Amount','Amount Received',
                        'Final Classification','Exception Reason',
                        'Notes','Email','Commercial']
        dateColumns = ['Disbersement Date', 'Payout Date', 'Payout Month']
        for col in dateColumns:
            clawback2[col] = clawback2[col].dt.date
        #book = xl.load_workbook(newMonthFilePath)
        writer = pd.ExcelWriter(newMonthFilePath, engine='openpyxl', mode='a', if_sheet_exists='overlay')
        clawback2[masterColumns].to_excel(writer, sheet_name=masterSheetName, header=None, index=False,
                startcol=1, startrow=masterMaxRow)
        writer.close()
        newMasterDf = pd.read_excel(newMonthFilePath, sheet_name = masterSheetName).iloc[:,1:]
        newMasterDf.rename(columns = {"Reserve Amount ":"Reserve Amount"}, inplace = True)
        newMasterDf['Exception Amount'] = newMasterDf['Exception %']*newMasterDf['Reserve Amount']
        newMasterDf['Clawback Amount'] = newMasterDf['Reserve Amount']-newMasterDf['Exception Amount']
        newMasterDf['Amount Received']=newMasterDf['Clawback Amount']
        logMessage = "New Clawback Masterfile for the month generated successfully."
        write_to_logger(logMessage)

              

        try:
            logMessage = "Refreshing created Clawback Masterfile"
            write_to_logger(logMessage)
            os.system('taskkill /f /im excel.exe')
            xlapp = client.DispatchEx("Excel.Application")
            wb = xlapp.Workbooks.Open(newMonthFilePath)
            wb.RefreshAll()
            xlapp.CalculateUntilAsyncQueriesDone()
            wb.Save()
            wb.Close(SaveChanges=True)
            xlapp.Quit()
            logMessage = "Clawback Masterfile refreshed successfully"
            write_to_logger(logMessage)
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            raise_fn_exception('write_to_master', e, exc_tb.tb_lineno)

        return newMasterDf
        
    except Exception as h:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('write_to_master', h, exc_tb.tb_lineno)

def update_gamers_sheet(gamersDf:pd.DataFrame, newMonthFilePath:str):
    # Load the workbook
    gamersSheetName = "Gamers"
    try:
        workbook = openpyxl.load_workbook(newMonthFilePath)

        # Get the sheet you want to clear (e.g., "Raw Data")
        sheet = workbook[gamersSheetName]

        # Delete rows and columns (e.g., clear 6 columns and 100 rows)
        #sheet.delete_cols(1, 6)
        sheet.delete_rows(2, 10000)
        workbook.save(newMonthFilePath)
        workbook.close()

        writer = pd.ExcelWriter(newMonthFilePath, engine='openpyxl', mode='a', if_sheet_exists='overlay')
        gamersDf.to_excel(writer, sheet_name=gamersSheetName, header=None, index=False,
                    startcol=0, startrow=1)
        writer.close()
        logMessage = "Gamers List updated in New Clawback Masterfile successfully."
        write_to_logger(logMessage)
    except Exception as h:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('update_gamers_sheet', h, exc_tb.tb_lineno)


def generate_trend_report(masterDf:pd.DataFrame, trendCsv:str):
    """Generates the summary report of the 'newMasterDf' dataframe and saves to the ClawbackOutFolder as given by 'trendCsv'\n
    param masterDf: type(pd.DataFrame). Dataframe of te master file to summarise.
    param trendCsv: type(str). Pathname of the trend file exported"""
    try:
        masterDf["Amount Received"] = masterDf["Amount Received"].astype(float)

        eligible_etf = masterDf[masterDf["Letter Version"]=='EFT'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        eligible_etf["Paid Out Loans"] = "Loans Eligible for Clawback (EFT)"

        eligible_qc_etf = masterDf[masterDf["Letter Version"]=='QC EFT'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        eligible_qc_etf["Paid Out Loans"] = "Loans Eligible for Clawback (QC EFT)"

        eligible_cheque = masterDf[masterDf["Letter Version"]=='Cheque'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        eligible_cheque["Paid Out Loans"] = "Loans Eligible for Clawback (Cheque)"

        eligible_amount = masterDf.pivot_table(columns = "Payout Month", values = "Reserve Amount",aggfunc = "sum" )
        eligible_amount["Paid Out Loans"] = "Total Reserve Eligible"

        exception_relationship = masterDf[masterDf["Exception Reason"]=='Waived for Relationship'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        exception_relationship["Paid Out Loans"] = "Exception Waived for Relationship"

        exception_insurance = masterDf[masterDf["Exception Reason"]=='Waived for Insurance Write Off'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        exception_insurance["Paid Out Loans"] = "Exception Waived for Insurance Write Off"

        exception_settled = masterDf[masterDf["Exception Reason"]=='Settled on less amount'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        exception_settled["Paid Out Loans"] = "Exception Settled on less amount"

        exception_branch = masterDf[masterDf["Exception Reason"]=='Branch Solicitation'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        exception_branch["Paid Out Loans"] = "Exception Branch Solicitation"

        exception_less_180 = masterDf[masterDf["Exception Reason"]=='Less then 180 days since booked'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        exception_less_180["Paid Out Loans"] = "Exception Less then 180 days since booked"

        exception_other = masterDf[masterDf["Exception Reason"]=='Other (rebooks, 1-off situations)'].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        exception_other["Paid Out Loans"] = "Exception Other (rebooks, 1-off situations)"

        exception_total = masterDf[~(masterDf["Exception Reason"]).isnull()].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        exception_total["Paid Out Loans"] = "Total Exceptions"

        exception_total_reserve = masterDf[(masterDf["Final Classification"])=="Exception "].pivot_table(columns = "Payout Month", values = "Exception Amount",aggfunc = "sum" )
        exception_total_reserve["Paid Out Loans"] = "Total Reserve Exceptions"

        loans_clawed_back_eft = masterDf[(masterDf["Letter Version"]=='EFT') & (masterDf["Amount Received"]>0)].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        loans_clawed_back_eft["Paid Out Loans"] = "Loans Clawed Back (EFT)"

        loans_clawed_back_qceft = masterDf[(masterDf["Letter Version"]=='QC EFT') & (masterDf["Amount Received"]>0)].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        loans_clawed_back_qceft["Paid Out Loans"] = "Loans Clawed Back (QC EFT)"

        loans_clawed_back_cheque = masterDf[(masterDf["Letter Version"]=='Cheque') & (masterDf["Amount Received"]>0)].pivot_table(columns = "Payout Month", values = "CIBC Client ID",aggfunc = "count" )
        loans_clawed_back_cheque["Paid Out Loans"] = "Loans Clawed Back (Cheque)"

        loans_clawed_back_total = masterDf[(masterDf["Amount Received"]>0)].pivot_table(columns = "Payout Month", values = "Amount Received",aggfunc = "sum" )
        loans_clawed_back_total["Paid Out Loans"] = "Total Reserve Clawed Back"

        trednDf=pd.concat([eligible_etf, eligible_qc_etf,
                eligible_cheque, eligible_amount,
                exception_relationship, exception_insurance,
                exception_settled, exception_branch,
                exception_less_180,
                exception_other, exception_total,
                exception_total_reserve,
                loans_clawed_back_eft,
                loans_clawed_back_qceft,
                loans_clawed_back_cheque,
                loans_clawed_back_total
                ], axis =0)

        percent_loans_clawed = pd.DataFrame(trednDf.iloc[-4:-1,].sum(axis=0)).transpose().drop("Paid Out Loans", axis = 1) / pd.DataFrame(trednDf.iloc[0:3,].sum(axis=0)).transpose().drop("Paid Out Loans", axis = 1)
        percent_loans_clawed["Paid Out Loans"] = "% of Eligible Loans Clawed Back"

        percent_reserve_clawed = trednDf.iloc[-1,].drop("Paid Out Loans") / trednDf.iloc[3,].drop("Paid Out Loans")
        percent_reserve_clawed["Paid Out Loans"] = "% of Eligible Reserve Clawed Back"
        percent_reserve_clawed = pd.DataFrame(percent_reserve_clawed).transpose()

        percent_exception = trednDf.iloc[11,].drop("Paid Out Loans") / trednDf.iloc[3,].drop("Paid Out Loans")
        percent_exception["Paid Out Loans"] = "Exception %"
        percent_exception = pd.DataFrame(percent_exception).transpose()

        trednDf=pd.concat([trednDf, percent_loans_clawed,percent_reserve_clawed, percent_exception], axis = 0)

        trednDf.set_index("Paid Out Loans", inplace = True, drop = True)

        trednDf.to_excel(trendCsv)
        logMessage = "Trend Report generated successfully"
        write_to_logger(logMessage)

        return trednDf
    except Exception as i:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('generate_trend_report', i, exc_tb.tb_lineno)  

def salesforce_email(clawback2:pd.DataFrame, sfMasterFile:pd.DataFrame, innovatecDealers:pd.DataFrame, dt:datetime.datetime, sfEmailCsv:str):
    """Generates the Salesforce Email file 'sfEmail4' to be sent to CIBC using the clawback2 dataframe and Salesforce email masterfile dataframe 'sfMasterFile'. \n
    Saves the email file to ClawbackOutFolder as sfEmailCsv.\n
    param clawback2: type(pd.DataFrame). Dataframe of clawbacks for the month.
    param sfMasterFile: type(pd.DataFrame). Dataframe of dealers info. Used as supplementary data source.
    param innovatecDealers: type(pd.DataFrame). Dataframe of dealers info from InnovatecCIBC. Used as a supplementary data source.
    param dt: type(datetime.datetime). Clawback month in view.
    param sfEmailCsv: type(str). Pathname of the Salesforce Email file exported.
    """

    sfcolumnMapper = {'Dealer':'Dealer’s Name',
                    'Dealer Street':'Street Address',
                    'Dealer City':'City',
                    'Dealer State/Province':'Province',
                    'Dealer Zip/Postal Code':'Postal Code',
                    'Reserve Clawback Email':'Clawback Email'}

    provinceReplaceMapper = {'Alberta':'AB',
                            'British Columbia':'BC',
                            'Manitoba':'MB',
                            'New Brunswick':'NB',
                            'Newfoundland and Labrador':'NL',
                            'Northwest Territories':'NT',
                            'Nova Scotia':'NS',
                            'Nunavut':'NU',
                            'Ontario':'ON',
                            'Prince Edward Island':'PE',
                            'Quebec':'QC',
                            'Saskatchewan':'SK',
                            'Yukon':'YT'}

    
    try:
        clawback2['RN'] = clawback2.sort_values(['Disbersement Date','Payout Date'], ascending=[True,True]) \
                .groupby(['Dealer ID']) \
                .cumcount() + 1
        
        clawback2.rename(columns={'CIBC Sales Rep':'DM',
                                    'CIBC Client ID':'Client',
                                     'PortalID':'Portal Id',
                                     'Vehicle VIN#':'Vehicle VIN',
                                     'Disbersement Date':'Date of Advance',
                                     'Payout Date':'Payout Date',
                                     'Reserve Amount':'Originations Fees Paid'
                                    }, inplace = True)
        
        firstRows = sorted(clawback2['Dealer ID'].unique())
        rnkColumns = sorted(clawback2['RN'].unique())
        subColumns = ['Client','Portal Id','Vehicle VIN','Date of Advance','Payout Date', 'Originations Fees Paid']
        sfEmail = pd.DataFrame(index = firstRows, columns = pd.MultiIndex.from_product([rnkColumns, subColumns],sortorder=0))
        
        sfEmail2 = sfEmail.reset_index(level = 0)
        sfEmail2.rename(columns = {"index":"Dealer ID"}, inplace = True)

        for i in range(0, len(sfEmail2)):
            for j in rnkColumns:
                for k in subColumns:
                    try:
                        sfEmail2.loc[i, (j, k)] = clawback2[(clawback2["Dealer ID"] == sfEmail2.loc[i,"Dealer ID"][0]) & (clawback2["RN"] ==j)][k].values[0]
                    except:
                        break
        
        #sfEmail2.to_csv(folderPath+"sfEmail2.csv")

        clawbackRnk1 = clawback2[clawback2["RN"]==1]
        clawbackRnk1['Contact Last Name'] = clawbackRnk1['Dealer']
        

        clawbackRnk1.set_index('Dealer ID', inplace = True)
        clawbackRnk1.columns = pd.MultiIndex.from_product([[''], list(clawbackRnk1.columns)])
        clawbackRnk1.reset_index(level = 0, inplace = True)
        

        
        sfMasterFile.set_index('Dealer ID', inplace = True)
        sfMasterFile.columns = pd.MultiIndex.from_product([[''], list(sfMasterFile.columns)])
        sfMasterFile.reset_index(level = 0, inplace = True)

        innovatecDealers.set_index('Dealer ID', inplace = True)
        innovatecDealers.columns = pd.MultiIndex.from_product([[''], list(innovatecDealers.columns)])
        innovatecDealers.reset_index(level = 0, inplace = True)

        sfEmail3 = sfEmail2.merge(clawbackRnk1, how = 'left', on = ["Dealer ID"])
        sfEmail4 = sfEmail3.merge(sfMasterFile, how = 'left', on = ["Dealer ID"])
        sfEmail4 = sfEmail4.merge(innovatecDealers, how = 'left', on = ["Dealer ID"])

        for i in range(0, len(sfEmail4)):
            if type(sfEmail4.loc[i, ('', 'Dealer Street')]) == type(None):
                if type(sfEmail4.loc[i, ('', 'Innovatec Street Address')]) == type(None):
                    sfEmail4.loc[i, ('', 'Dealer Street')] = sfEmail4.loc[i, ('', 'Street Address')]
                else:
                    sfEmail4.loc[i, ('', 'Dealer Street')] = sfEmail4.loc[i, ('', 'Innovatec Street Address')]
            if type(sfEmail4.loc[i, ('', 'Dealer City')]) == type(None):
                if type(sfEmail4.loc[i, ('', 'Innovatec City')]) == type(None):
                    sfEmail4.loc[i, ('', 'Dealer City')] = sfEmail4.loc[i, ('', 'City')]
                else:
                    sfEmail4.loc[i, ('', 'Dealer City')] = sfEmail4.loc[i, ('', 'Innovatec City')]
            if type(sfEmail4.loc[i, ('', 'Dealer State/Province')]) == type(None):
                if type(sfEmail4.loc[i, ('', 'Innovatec Province')]) == type(None):
                    sfEmail4.loc[i, ('', 'Dealer State/Province')] = sfEmail4.loc[i, ('', 'Province')] 
                else:
                    sfEmail4.loc[i, ('', 'Dealer State/Province')] = sfEmail4.loc[i, ('', 'Innovatec Province')] 
            if type(sfEmail4.loc[i, ('', 'Dealer Zip/Postal Code')]) == type(None):
                if type(sfEmail4.loc[i, ('', 'Innovatec Postal Code')]) ==type(None):
                    sfEmail4.loc[i, ('', 'Dealer Zip/Postal Code')] = sfEmail4.loc[i, ('', 'Postal Code')]
                else:
                    sfEmail4.loc[i, ('', 'Dealer Zip/Postal Code')] = sfEmail4.loc[i, ('', 'Innovatec Postal Code')]
            if type(sfEmail4.loc[i, ('', 'Reserve Clawback Email')]) == type(None):
                if type(sfEmail4.loc[i, ('', 'Innovatec EFTEmail')]) ==type(None):
                    sfEmail4.loc[i, ('', 'Reserve Clawback Email')] = sfEmail4.loc[i, ('', 'Clawback Email')]
                else:
                    sfEmail4.loc[i, ('', 'Reserve Clawback Email')] = sfEmail4.loc[i, ('', 'Innovatec EFTEmail')]
        
        sfEmail4[('', 'Date sending [Month Day, Year]')] = None
        sfEmail4[('', 'Contact First Name')] = 'Clawback'
        sfEmail4[('', 'Pay back no later then [Month Day, Year]')] = None
        sfEmail4[('', ' ')] = None
        emailColumns = [('Dealer ID', ''),
            ('', 'Letter Version'), 
            ('', 'Date sending [Month Day, Year]'),
            ('', 'Dealer'),
            ('', 'Dealer Street'),
            ('', 'Dealer City'),
            ('', 'Dealer State/Province'),
            ('', 'Dealer Zip/Postal Code'),
            ('', 'Dealer Contact'),
            ('', 'Contact First Name'),
            ('', 'Contact Last Name'),
            ('', 'Reserve Clawback Email'),
            ('', 'Pay back no later then [Month Day, Year]'),
            ('', 'DM'),
            ('', ' ')]
        for x in list(sfEmail2.columns)[1:]:
            emailColumns.append(x)
            
        sfEmail4 = sfEmail4[emailColumns]   

        searchword = ['Disbersement Date\')', 'Payout Date\')', 'Date of Advance\')']
        searchFiles = []
        for word in searchword:
            for item in (list(sfEmail4.columns)):
                if str(item).endswith(word):
                    searchFiles.append(item)        
        
        for col in searchFiles:
            sfEmail4[col] = sfEmail4[col].astype('<M8[ns]')
            sfEmail4[col] = sfEmail4[col].dt.strftime('%m/%d/%Y')

        newColumns = []
        for i in range(0,len(sfEmail4.columns)):
            if i >=14:
                newColumns.append((sfEmail4.columns[i][1]+"#"+str(sfEmail4.columns[i][0])).strip())
                #newColumns.append((sfEmail4.columns[i][0], (sfEmail4.columns[i][1]+"#"+str(sfEmail4.columns[i][0])).strip()))
            else:
                newColumns.append((sfEmail4.columns[i][1]+str(sfEmail4.columns[i][0])).strip())
                ##newColumns.append((sfEmail4.columns[i][0], (sfEmail4.columns[i][1]+str(sfEmail4.columns[i][0])).strip()))
        #sfEmail4.columns = pd.MultiIndex.from_tuples(newColumns)
        sfEmail4.columns = newColumns
        sfEmail4['Dealer State/Province'].replace(provinceReplaceMapper, inplace = True)
        sfEmail4.rename(columns = sfcolumnMapper, inplace = True)
        
        sfEmail4['Dealer Contact'] = sfEmail4['Province'].apply(lambda x: "A qui de droip:" if x=='QC' else "To whom it may concern")
        sfEmail4['Date sending [Month Day, Year]'] = sfEmail4['Province'].apply(lambda x: date_to_locale(datetime.date.today(), x))
        sfEmail4['Pay back no later then [Month Day, Year]'] = sfEmail4['Province'].apply(lambda x: date_to_locale(get_last_working_date((dt + pd.tseries.offsets.MonthEnd(2))),x)) ###last business day
        
        
        
        sfEmail4.to_excel(sfEmailCsv, index = False) ###encoding = "windows-1252"
        logMessage = "Salesforce Email file generated successfully."
        write_to_logger(logMessage)

        return sfEmail4
    except Exception as j:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('salesforce_email', j, exc_tb.tb_lineno)

def select_clawback_files():
    clawbackFile = None
    masterFile = None
    clawbackFile = file_select_window("clawback file")
    if clawbackFile == None or len(clawbackFile)==0:
        return None, None
    filePath = clawbackFile.title()
    logMessage = "Clawback File selected: "+ filePath
    write_to_logger(logMessage)

    masterFile = file_select_window("master file")
    if masterFile == None or len(masterFile)==0:
        return None, None
    masterFilePath = masterFile.title()
    logMessage = "Previous month Master File selected: "+ masterFilePath
    write_to_logger(logMessage)

    return filePath, masterFilePath

def generate_clawback_file():
    """Executes the Initial Clawback Generation process using GUI to select the workfile for the month and the masterfile for the previous month"""

    folderPath = ClawbackOutFolder
    masterSheetName = "ELIGIBLE CLAWBCKS"

    filePath, masterFilePath = select_clawback_files()

    if type(filePath) == type(None) or type(masterFilePath) == type(None):
        return False

    df, gamersDf, dealerBody, masterDf, eftChequeDf, cibcDf, sfMasterFile, innovatecDealers, dt = initiate_clawback_files(filePath, masterFilePath)

    fileDate = dt.strftime('%Y_%m_%d')
    fileDateString = datetime.datetime.strptime(fileDate,'%Y_%m_%d').strftime('%B, %Y')

    file = filePath.split('/')[-1]    
    masterFile = masterFilePath.split('/')[-1]

    ##Outfiles
    clawbackCsv = folderPath+"Eligible_Clawback_"+fileDate+".xlsx"
    trendCsv = folderPath+"Clawback_Summary_Trend_"+fileDate+".xlsx"
    sfEmailCsv = folderPath+"Salesforce_Email_"+fileDate+".xlsx"
    exceptionsCsv = folderPath+"Exceptions_List_"+fileDate+".xlsx"
    newMonthFilePath = folderPath+ "CIBC RESERVE CLAWBACK MASTER "+fileDateString+".xlsx"
    
    try:
        exceptionsDf = exceptions_master(masterDf, masterSheetName, exceptionsCsv)        
        clawback2 = eligible_clawback(df,cibcDf,dealerBody,eftChequeDf,gamersDf,exceptionsDf, dt, clawbackCsv)           
        masterMaxRow = len(masterDf)+1        
        newMasterDf = write_to_master( masterFilePath, newMonthFilePath, masterSheetName, clawback2, masterMaxRow)        
        trendDf = generate_trend_report(newMasterDf,trendCsv)        
        sfEmailDf = salesforce_email(clawback2, sfMasterFile, innovatecDealers, dt, sfEmailCsv)
        update_gamers_sheet(gamersDf, newMonthFilePath)
        process_completed("clawback")        
        write_to_logger("Generate Clawback File completed")
        #return clawback2,trendDf,sfEmailDf, exceptionsDf, newMasterDf, True
        return True
    
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('main', e, exc_tb.tb_lineno)


def generate_ach_file():
    """Executes the Final Clawback and ACH File Generation process."""
    date, masterFile = clawback_for_ach()
    if type(date) == type(None) or type(masterFile) == type(None):
        return False
    eligibleClawback, monthExceptions, clawback2 = final_eligible_clawback(date,masterFile)
    achExceptions = get_ach_exceptions(eligibleClawback)
    if type(achExceptions) == str:
        if achExceptions == 'timeout':
            process_completed("timeout")
    elif len(achExceptions) >0:
        exceptionsCnt = len(achExceptions)
        process_completed("achException", exceptionsCnt)
    else:
        process_completed("noException")
    return True


def program_select_window():
    """Creates a GUI for users to select the Clawback process they wish to initiate: \n
    1. Generate Clawback File - the initial clawback process. Executed with the receipt of the Clawback source file from CIBC for the month.\n
    2. Generate ACH File - the final clawback process. Executed at the end of the month when exceptions are collated and sent from Jonathan"""

    
    layoutText = "Please select a Process"
    layout = [
        [sg.Text(layoutText)],
        [sg.Button("Generate Clawback Files", size=(10,3)), sg.Button("Generate ACH Files", size=(10,3))]
       
    ]
    
    window = sg.Window("Choose Process", layout, finalize=True, size = (225,120), element_justification = 'center')
    
    while True:
        event, values = window.read()
        if event == "Generate Clawback Files":
            logMessage = "Generate Clawback Files selected"
            write_to_logger(logMessage)            
            status = generate_clawback_file()
            if status:
                break
            else:
                logMessage = "Generate Clawback Files cancelled"
                write_to_logger(logMessage)
                window.close()
                program_select_window()
            break

        elif event == "Generate ACH Files":            
            logMessage = "Generate ACH Files selected"
            write_to_logger(logMessage)
            status = generate_ach_file()
            if status:
                break
            else:
                logMessage = "Generate ACH Files cancelled"
                write_to_logger(logMessage)
                window.close()
                program_select_window()


        elif event == sg.WIN_CLOSED:
            break

    logMessage = "Exiting Clawback Window"
    write_to_logger(logMessage)
    window.close()
    return 

def main():
    """
    The main function calls the program_select_window function and handles any exceptions that occur.
    """
    try:
        program_select_window()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('main', e, exc_tb.tb_lineno)

if __name__ == '__main__':
    try:
        warnings.filterwarnings('ignore')
        logFile = start_logger()
        main()
        logFile.close()
        sys.exit()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise_fn_exception('clawback_with_UI', e, exc_tb.tb_lineno)
