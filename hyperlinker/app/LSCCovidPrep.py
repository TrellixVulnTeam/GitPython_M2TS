from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/LSCCovidPrep", methods=['GET', 'POST'])
def LSCCovidPrep():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        test = pd.read_excel(f)
        
        test.fillna('',inplace=True)
        
        #Cleaning
        if test.iloc[0][0] == '':
            df = pd.read_excel(f,skiprows=2)
        else:
            df = pd.read_excel(f)
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)          
        
        
        #Funding Source: 001 is 4040 (primary or secondary), 002 is 4000 (primary or secondary), else 003

        def FundingSourceCode(PrimaryFunding,SecondaryFunding):
            if PrimaryFunding.startswith("4040") == True or SecondaryFunding.startswith("4040") == True:
                return "001"
            elif PrimaryFunding.startswith("4000") == True or SecondaryFunding.startswith("4000") == True:
                return "002"
            else:  
                return "003"
        
        df['Funding Source Code'] = df.apply(lambda x: FundingSourceCode(x['Primary Funding Codes'],x['Secondary Funding Codes']), axis = 1)
        
        #Staffing Code: S is staff, P is Private Attorney Involvement
        
        def StaffCode(PAI,PBI):
            if PAI == "Yes":
                return "P"
            elif PBI == "Yes":
                return "P"
            else:
                return "S"
        
        df['Staffing Code'] = df.apply(lambda x: StaffCode(x['PAI Case?'],x['Matter has current Pro Bono involvement']), axis = 1)
        
        
        #Problem Code: first 2 digits from LPC
        df['Problem Code'] = df['Legal Problem Code'].str.slice(0,2)
        
        #Closing Code: first letter or two letters from closing code, or blank if open
        
        def ClosingCode(CloseReason):
            if CloseReason == "":
                return ""
            elif CloseReason.startswith('IA') == True:
                return "Ia"
            elif CloseReason.startswith('IB') == True:
                return "Ib"
            elif CloseReason.startswith('IC') == True:
                return "Ic"
            else:
                return CloseReason[0:1]
        df['Closing Code'] = df.apply(lambda x: ClosingCode(x['Close Reason']), axis = 1)

        
        #Gender Code: MWOU
        def LSCGender(Gender):
            if Gender.startswith("M") == True or Gender.startswith("Transgender M"):
                return "M"
            elif Gender.startswith("F") == True or Gender.startswith("Transgender W") == True:
                return "W"
            elif Gender.startswith("O") == True or Gender.startswith("S") == True:
                return "O"
            else:
                return "U"
        df['Gender Code'] = df.apply(lambda x: LSCGender(x['Gender']), axis = 1)
        
        #Veteran Status Code: N onveteran V eteran U nknown
        def VeteranCode(Veteran):
            if Veteran == "No":
                return "N"
            elif Veteran == "Yes":
                return "V"
            else: 
                return "U"
        
        df['Veteran Status Code'] = df.apply(lambda x: VeteranCode(x['Veteran']), axis = 1)
        
        #Ethnicity Code: A sian B lack H ispanic N ativeamerican O ther W hite U nknown
        def EthnicityCode(Race):
            if Race == "Hispanic" or Race == "Latina/o":
                return "H"
            elif Race == "Asian or Pacific Islander" or Race == "South Asian":
                return "A"
            elif Race == "White (Not Hispanic)":
                return "W"
            elif Race == "Self-Identified/Other":
                return "O"
            elif Race == "Black/African American/African Descent":
                return "B"
            elif Race == "Prefer Not To Say":
                return "U"
            elif Race == "Native American/American Indian":
                return "N"
            else:
                return "U"
        
        df['Ethnicity Code'] = df.apply(lambda x: EthnicityCode(x['Race']), axis = 1)
        
        #Year of Birth - client
        df['Year of Birth'] = df['Date of Birth'].str.slice(0,4)
        
        df['Year of Birth'] = df['Year of Birth'].apply(lambda x: '1900' if x == '' else x)
        
        
        df['ClosedDateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Closed']), axis=1)
        df['OpenedDateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']), axis=1)
        
        
        #Closed or Opened in Prior Quarter ** NEED To UpDate this every quarter? BAD!***
        def DateInQuarter(Date):
            if Date == "":
                return ""
            
            elif Date >= 20210401 and Date <= 20210630:
                return "Yes"
            else:
                return "No"
        
        df['ClosedInReportingQuarter'] = df.apply(lambda x: DateInQuarter(x['ClosedDateConstruct']), axis = 1)
       
        df['OpenedInReportingQuarter'] = df.apply(lambda x: DateInQuarter(x['OpenedDateConstruct']), axis = 1)
        
        def CSRDeterminer(LSCEligible,AssistanceDocumented,TimelyClosing,CitizenshipStatus,AttestationOnFile,VerifiedNonCitizenship,ClientInPerson,CloseReason,LegalServerCSR):
            if CloseReason == '':
                return LegalServerCSR
            elif LSCEligible != 'Yes':
                return 'No'
            elif AssistanceDocumented != 'Yes':
                return 'No'
            elif TimelyClosing != 'Yes':
                return 'No'
            else:
                if CitizenshipStatus == 'Citizen' and AttestationOnFile == 'Yes':
                    return 'Yes'
                elif CitizenshipStatus == 'Non-Citizen' and VerifiedNonCitizenship == 'Yes':
                    return 'Yes'
                elif CloseReason.startswith(('A','B')) == True and ClientInPerson == 'No':
                    return 'Yes'
                else:
                    return 'No'
                    
        df['Python CSR Status'] = df.apply(lambda x: CSRDeterminer(x['LSC Eligible?'],x['CSR: Is Legal Assistance Documented?'],x['CSR: Timely Closing?'],x['Citizenship Status'],x['Attestation on File?'],x['Staff Verified Non-Citizenship Documentation'],x['Did any Staff Meet Client in Person?'],x['Close Reason'],x['CSR Eligible']), axis = 1)
        
        
        
        #Should a case be included in the report or deleted?
        #All Cases must be LSC/CSR Eligible (CSR can't be no, but only has to be yes [non-blank] if it's closed)
        #All Cases must be Covid-Related
        #From what remains, just report cases Opened or Closed in reporting period
        
        def IncludeInReport(InvolvesCovid,LSCElig,PythonCSRElig,DateClosed,OpenedInReportingQuarter,ClosedInReportingQuarter):
            if InvolvesCovid != "Yes":
                return "Remove"
            elif LSCElig != "Yes":
                return "Remove"
            elif DateClosed != "" and PythonCSRElig != "Yes":
                return "Remove"
            elif PythonCSRElig == "No":
                return "Remove"
            elif OpenedInReportingQuarter == "Yes" or ClosedInReportingQuarter == "Yes":
                return "Include"
            

        df['IncludeInReport?'] = df.apply(lambda x: IncludeInReport(x['Case Involves Covid-19'],x['LSC Eligible?'],x['Python CSR Status'],x['Date Closed'],x['OpenedInReportingQuarter'],x['ClosedInReportingQuarter']), axis = 1)

        #Put everything in the right order
        
        #this is just so that it doesn't split the spreadsheets:
        df['PlaceholderColumn'] = '!'
        
        df = df[[
        "Funding Source Code",
        "Staffing Code",
        "Problem Code",
        "Closing Code",
        "Gender Code",
        "Veteran Status Code",
        "Ethnicity Code",
        "Year of Birth",
        "PlaceholderColumn",
        'Hyperlinked CaseID#',
        'Primary Advocate',
        "Date Opened",
        "OpenedDateConstruct",
        "Date Closed",
        "ClosedDateConstruct",
        "Legal Problem Code",
        'Primary Funding Codes',
        'Secondary Funding Codes',
        'Case Involves Covid-19',
        'OpenedInReportingQuarter',
        'ClosedInReportingQuarter',
        'CSR Eligible',
        'Python CSR Status',
        'LSC Eligible?',
        'IncludeInReport?'
        ]]      
        
        
        
        #Preparing Excel Document
        
        #Split into different tabs
        allgood_dictionary = dict(tuple(df.groupby('IncludeInReport?')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                ws = writer.sheets[i]
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                ws.set_column('J:J',20,link_format)
                ws.set_column('A:ZZ',25)
                ws.freeze_panes(1,0)
                            
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename)
       
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Prepared " + f.filename)

    return '''
    <!doctype html>
    <title>LSC Covid Report Prep</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>Prepare Dataset for LSC Covid Report</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Prepare!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with: <a href="https://lsnyc.legalserver.org/report/dynamic?load=1725" target="_blank">Pascale Big Base Report</a>.</li>
    
    </br>
    <a href="/">Home</a>
    '''
