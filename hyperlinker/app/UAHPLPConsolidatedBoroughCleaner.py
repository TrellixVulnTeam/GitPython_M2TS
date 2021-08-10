from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/UAHPLPConsolidatedBoroughCleaner", methods=['GET', 'POST'])
def UAHPLPConsolidatedBoroughCleaner():
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


        
        
        #Has to have a Housing Type of Case
        

        #Has to have a Housing Level of Service 

        #Level of Service is HOLD FOR REVIEW tester? incorporate into above

        #Eligiblity date tester - blank or not?
       
        df['EligDateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        df['Pre-3/1/20 Elig Date?'] = df.apply(lambda x: HousingToolBox.PreThreeOne(x['EligDateConstruct']), axis=1)
        
        df['OpenedDateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']), axis=1)
       
        #don't need consent form if it's post-covid advice/brief (still need release if we do have the index number)
        def ReleaseTester(HRARelease,LevelOfService,IndexNum):
            LevelOfService = str(LevelOfService)
            IndexNum = str(IndexNum)
            if LevelOfService.startswith("Advice") or LevelOfService.startswith("Brief"):
        
                if IndexNum == '' or IndexNum.startswith('N') == True or IndexNum.startswith('n') == True:
                    return 'Unnecessary due to Limited Service'
                else:
                    return HRARelease
            else:
                return HRARelease
       
        df['HRA Release?'] = df.apply(lambda x: ReleaseTester(x['HRA Release?'],x["Housing Level of Service"],x['Gen Case Index Number']),axis = 1)
       
        #PA Tester if theres no dhci, not needed for post-covid advice/brief cases
        def PATester (PANum,DHCI,PreThreeOne,LevelOfService,EligDate,OpenDate):
            LevelOfService = str(LevelOfService)
            if PreThreeOne == "No" and LevelOfService.startswith("Advice") and EligDate != '':
                return "Unnecessary due limited service"
            elif PreThreeOne == "No" and LevelOfService.startswith("Advice") and OpenDate >= 20200301:
                return "Unnecessary due to limited service"
            
            elif DHCI == "DHCI Form" and PANum == "":
                return "Not Needed due to DHCI"
            else:
                return PANum
            
        df['Gen Pub Assist Case Number'] = df.apply(lambda x: PATester(x['Gen Pub Assist Case Number'],x['Housing Income Verification'],x['Pre-3/1/20 Elig Date?'],x["Housing Level of Service"],x['EligDateConstruct'],x['OpenedDateConstruct']),axis = 1)
        
        #Outcome Tester - date no outcome or outcome no date
        
        def OutcomeTester(Outcome,OutcomeDate):
            if OutcomeDate != "" and Outcome == "":
                return "Needs Outcome"
            else:
                return Outcome
                
        df['Housing Outcome'] = df.apply(lambda x: OutcomeTester(x['Housing Outcome'],x['Housing Outcome Date']),axis = 1)
                
        def OutcomeDateTester(Outcome,OutcomeDate):
            if Outcome != "" and OutcomeDate == "":
                return "Needs Outcome Date"
            else:
                return OutcomeDate      

        df['Housing Outcome Date'] = df.apply(lambda x: OutcomeDateTester(x['Housing Outcome'],x['Housing Outcome Date']),axis = 1)
        
        
        #Test if a case is actually a housing case
        def NonHousingCase (LegalProblemCode):
            if LegalProblemCode.startswith('6') == True:
                return ''
            else:
                return 'Non-Housing Case, Please Review'
                
        df['Housing Case?'] = df.apply(lambda x: NonHousingCase(x['Legal Problem Code']),axis = 1)
            
    
        #Is everything okay with a case? Also remove if Eligdate is from prior year

        def TesterTester (EligConstruct,HRARelease,HousingLevel,HousingType,EligDate,PANum,Outcome,OutcomeDate,HousingCase):
           
            if EligConstruct != '' and EligConstruct < 20200701 :
                return 'Eligibility date from prior contract year'
            elif HRARelease == "" or HRARelease == "No" or HRARelease == " ":
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
            elif HousingCase == 'Non-Housing Case, Please Review':
                return 'Case Needs Attention'
            else:
                return 'No Cleanup Necessary'
            
        df['Tester Tester'] = df.apply(lambda x: TesterTester(x['EligDateConstruct'],x['HRA Release?'],x['Housing Level of Service'],x['Housing Type Of Case'],x['HAL Eligibility Date'],x['Gen Pub Assist Case Number'],x['Housing Outcome'],x['Housing Outcome Date'],x['Housing Case?']),axis=1)
        
        
        #Delete if everything's okay **

        df = df[df['Tester Tester'] == "Case Needs Attention"]

        #sort by case handler
        
        df = df.sort_values(by=['Primary Advocate'])
        
        
        
        #Put everything in the right order
        
        df = df[['Hyperlinked CaseID#',
        'Primary Advocate',
        'Date Opened',
        'Date Closed',
        "Client First Name",
        "Client Last Name",
        
        "HRA Release?",
        "Housing Level of Service",
        "Housing Type Of Case",
        "HAL Eligibility Date",
        
        
        
        "Gen Pub Assist Case Number",
        
        "Housing Income Verification",
        #"Housing Signed DHCI Form",
        
        "Housing Outcome",
        "Housing Outcome Date",
        "Housing Case?",
        "Gen Case Index Number", 
        "Tester Tester",
        "Assigned Branch/CC"
        
        ]]      
        
        #Preparing Excel Document
        
        #Split into different tabs
        allgood_dictionary = dict(tuple(df.groupby('Assigned Branch/CC')))
        
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

                ws.conditional_format('G2:K100000',{'type': 'blanks',
                                                 'format': problem_format})
                ws.conditional_format('H2:H100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Hold For Review',
                                                 'format': problem_format})
                ws.conditional_format('G2:G100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'No',
                                                 'format': problem_format})
                ws.conditional_format('M2:N100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
                ws.conditional_format('O2:O100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Review',
                                                 'format': problem_format})
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename)
       
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>UAHPLP Consolidated Borough Cleaner</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>UAHPLP Consolidated Borough Cleaner:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2308" target="_blank">HPLP/UAC Internal Report All Cases</a>.</li>
    
   
    </br>
    <a href="/">Home</a>
    '''
