from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/MLSIntakeConsolidatedHousingCleaner", methods=['GET', 'POST'])
def MLSIntakeConsolidatedHousingCleaner():
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
        
        df['Assigned Branch/CC'] = df.apply(lambda x : DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis=1)   


        #Delete cases closed before 7/1/2021
        
        df['ClosedConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Closed']), axis=1)
                
        def CaseSelector(DateClosed,EligDate):
            if EligDate == '':
                if DateClosed == '':
                    return 'Keep'
                elif DateClosed >= 20210701:
                    return 'Keep'
                else:
                    'Do not keep'
            else:
                return 'Keep'
        
        df['Keep?'] = df.apply(lambda x: CaseSelector(x['ClosedConstruct'],x['HAL Eligibility Date']),axis=1)
        
        df = df[df['Keep?'] == 'Keep']
        
       
        #Eligiblity date tester - blank or not?
       
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        df['Post 12/1/21 Elig Date?'] = df.apply(lambda x: HousingToolBox.PostTwelveOne(x['DateConstruct']), axis=1)
       
        #don't need consent form if it's post-covid advice/brief
        def ReleaseTester(HRARelease,LevelOfService,PostTwelveOne,PrimaryCode,DateConstruct):
            LevelOfService = str(LevelOfService)
            if DateConstruct == "":
                return HRARelease
            elif PostTwelveOne == "No":
                if LevelOfService.startswith("Advice"):
                    return "Unnecessary advice/brief"
                elif PrimaryCode.startswith("31") and LevelOfService.startswith("Brief"):
                    return "Unnecessary advice/brief"
                else:
                    return HRARelease
            else:
                return HRARelease
       
        df['HRA Release?'] = df.apply(lambda x: ReleaseTester(x['HRA Release?'],x["Housing Level of Service"],x['Post 12/1/21 Elig Date?'],x["Primary Funding Code"],x['DateConstruct']),axis = 1)
        
        df['Housing Income Verification'] = df.apply(lambda x: ReleaseTester(x['Housing Income Verification'],x["Housing Level of Service"],x['Post 12/1/21 Elig Date?'],x["Primary Funding Code"],x['DateConstruct']),axis = 1)
        
       
       
        #PA Tester if theres no dhci, not needed for post-covid advice/brief cases
        def PATester (PANum,DHCI,LevelOfService,DateConstruct,PostTwelveOne,PrimaryCode):
            LevelOfService = str(LevelOfService)
            if DateConstruct == "":
                if DHCI == "DHCI Form" and PANum == "":
                    return "Not Needed due to DHCI"
                else:
                    return PANum
            elif PostTwelveOne == "No":
                if LevelOfService.startswith("Advice"):
                    return "Unnecessary advice/brief"
                elif PrimaryCode.startswith("31") and LevelOfService.startswith("Brief"):
                    return "Unnecessary advice/brief"
                elif DHCI == "DHCI Form" and PANum == "":
                    return "Not Needed due to DHCI"
                else:
                    return PANum
            else:
                return PANum
            
        df['Gen Pub Assist Case Number'] = df.apply(lambda x: PATester(x['Gen Pub Assist Case Number'],x['Housing Income Verification'],x["Housing Level of Service"],x['DateConstruct'],x['Post 12/1/21 Elig Date?'],x["Primary Funding Code"]),axis = 1)

        #Outcome Tester - date no outcome or outcome no date
        
        def OutcomeTester(Outcome,OutcomeDate):
            if OutcomeDate != "" and Outcome == "":
                return "Needs Outcome"
            else:
                return Outcome
                
        def OutcomeDateTester(Outcome,OutcomeDate):
            if Outcome != "" and OutcomeDate == "":
                return "Needs Outcome Date"
            else:
                return OutcomeDate      
       
        df['Housing Outcome'] = df.apply(lambda x: OutcomeTester(x['Housing Outcome'],x['Housing Outcome Date']),axis = 1)
        df['Housing Outcome Date'] = df.apply(lambda x: OutcomeDateTester(x['Housing Outcome'],x['Housing Outcome Date']),axis = 1)
        
        
        
    
        #Is everything okay with a case? 

        def TesterTester (HRARelease,HousingLevel,HousingType,EligDate,PANum,Outcome,OutcomeDate):
           
            if HRARelease == "" or HRARelease == "No" or HRARelease == " ":
                return 'Case Needs Attention'
            elif HousingLevel == "" or HousingLevel == "Hold For Review":
                return 'Case Needs Attention'
            elif HousingType == '':
                return 'Case Needs Attention'
            elif EligDate == '':
                return 'Case Needs Attention'
            elif PANum == '':
                return 'Case Needs Attention'
            elif Outcome == 'Needs Outcome' or OutcomeDate == 'Needs Outcome Date':
                return 'Case Needs Attention'
            else:
                return 'No Cleanup Necessary'
            
        df['Tester Tester'] = df.apply(lambda x: TesterTester(x['HRA Release?'],x['Housing Level of Service'],x['Housing Type Of Case'],x['HAL Eligibility Date'],x['Gen Pub Assist Case Number'],x['Housing Outcome'],x['Housing Outcome Date']),axis=1)
        
        
        #Delete if everything's okay **

        #df = df[df['Tester Tester'] == "Case Needs Attention"]

        
        #assign casehandlers to Intake Paralegals:
        df['Assigned Paralegal'] = df.apply(lambda x: HousingToolBox.MLSIntakeAssign(x['Primary Advocate'],x['Caseworker Name']),axis = 1)

        #sort by case handler
        
        df = df.sort_values(by=['Primary Advocate'])
        
        
        
        #Put everything in the right order
        
        df = df[['Hyperlinked CaseID#',
        'Primary Advocate',
        'Date Opened',
        "Client First Name",
        "Client Last Name",
        
        "HRA Release?",
        "Housing Level of Service",
        "Housing Type Of Case",
        "HAL Eligibility Date",
        
        
        "Gen Case Index Number",

        "Gen Pub Assist Case Number",
        "Housing Income Verification",

        
        "Housing Outcome",
        "Housing Outcome Date",
        
        "Primary Funding Code",
        "Service Date",
        "Combined Notes",
        "Activity Details",
        "Total Time For Case",
        
        
        
        "Tester Tester",
        "Assigned Branch/CC",
        "Assigned Paralegal",
        "Caseworker Name",
        "Date Closed"
        ]]      
        
        #Preparing Excel Document
        
        #Split into different tabs
        allgood_dictionary = dict(tuple(df.groupby('Assigned Paralegal')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                ws = writer.sheets[i]
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                regular_format = workbook.add_format({'font_color':'black'})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
                ws.autofilter('B1:ZZ1')
                ws.freeze_panes(1, 2)

                FLRowRange='F1:L'+str(dict_df[i].shape[0]+1)
                print(FLRowRange)


                ws.conditional_format(FLRowRange,{'type': 'blanks',
                                                 'format': problem_format})
                ws.conditional_format('G2:G100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Hold For Review',
                                                 'format': problem_format})
                ws.conditional_format('F2:F100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'No',
                                                 'format': problem_format})
                ws.conditional_format('M2:N100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename)
       
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>MLS Housing Consolidated intake Cleaner</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>MLS Housing Consolidated intake Cleaner:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2315" target="_blank">MLS Consolidated Housing Cleanup Report</a>.</li>
    
   
    </br>
    <a href="/">Home</a>
    '''
