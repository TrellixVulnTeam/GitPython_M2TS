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


        #Has to have a Housing Type of Case
        

        #Has to have a Housing Level of Service 

       
        #Eligiblity date tester - blank or not?
       
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        df['Pre-3/1/20 Elig Date?'] = df.apply(lambda x: HousingToolBox.PreThreeOne(x['DateConstruct']), axis=1)
       
        #don't need consent form if it's post-covid advice/brief
        def ReleaseTester(HRARelease,PreThreeOne,LevelOfService):
            LevelOfService = str(LevelOfService)
            if PreThreeOne == "No" and LevelOfService.startswith("Advice"):
                return "Unnecessary advice/brief"
            else:
                return HRARelease
       
        df['HRA Release?'] = df.apply(lambda x: ReleaseTester(x['HRA Release?'],x['Pre-3/1/20 Elig Date?'],x["Housing Level of Service"]),axis = 1)
        
        df['Housing Income Verification'] = df.apply(lambda x: ReleaseTester(x['Housing Income Verification'],x['Pre-3/1/20 Elig Date?'],x["Housing Level of Service"]),axis = 1)
        
       
       
        #PA Tester if theres no dhci, not needed for post-covid advice/brief cases
        def PATester (PANum,DHCI,PreThreeOne,LevelOfService):
            LevelOfService = str(LevelOfService)
            if PreThreeOne == "No" and LevelOfService.startswith("Advice"):
                return "Unnecessary advice/brief"
            elif DHCI == "DHCI Form" and PANum == "":
                return "Not Needed due to DHCI"
            else:
                return PANum
            
        df['Gen Pub Assist Case Number'] = df.apply(lambda x: PATester(x['Gen Pub Assist Case Number'],x['Housing Income Verification'],x['Pre-3/1/20 Elig Date?'],x["Housing Level of Service"]),axis = 1)

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

        df = df[df['Tester Tester'] == "Case Needs Attention"]

        #assign casehandlers to Intake Paralegals:
        
        Evelyn_Casehandlers = ['Delgadillo, Omar','Heller, Steven E','Latterner, Matt J','Robles-Castillo, Camila J','Tilyayeva, Rakhil','Almanzar, Yocari']
        Diana_V_Casehandlers = ['Abbas, Sayeda','Evers, Erin C.','Hao, Lindsay','He, Ricky','Sharma, Sagar','Spencer, Eleanor G','Wilkes, Nicole','Allen, Sharette','Ortiz, Matthew B','Sun, Dao','Risener, Jennifer A','Evers, Erin C.','Surface, Ben L']
        Diana_G_Casehandlers = ['Frierson, Jerome C','Saxton, Jonathan G','Orsini, Mary K','Allen, Sharette','Duffy-Greaves, Kevin','Freeman, Daniel A','Gokhale, Aparna S','Gonzalez, Matias','Gonzalez, Matias G','Labossiere, Samantha J.','Shah, Ami Mahendra']
        Keiannis_Casehandlers = ['Almanzar, Milagros','Briggs, John M','Dittakavi, Archana','Gonzalez-Munoz, Rossana G','Honan, Thomas J','James, Lelia','Kelly, Kitanya','Mottley, Darlene','Yamasaki, Emily Woo J','McCune, Mary','Vogltanz, Amy K']
        Dennis_Casehandlers = ['Braudy, Erica','Kulig, Jessica M','Mercedes, Jannelys J','Harshberger, Sae','Black, Rosalind','Basu, Shantonu J']
        Rosa_Casehandlers = ['Acron, Denise D','Anunkor, Ifeoma O','Reyes, Nicole V','Vega, Rita']

        def IntakeAssign(Casehandler):
            if Casehandler in Evelyn_Casehandlers:
                return "Evelyn V."
            elif Casehandler in Diana_V_Casehandlers:
                return "Diana V."
            elif Casehandler in Diana_G_Casehandlers:
                return "Diana G."
            elif Casehandler in Keiannis_Casehandlers:
                return "Keiannis G."
            elif Casehandler in Dennis_Casehandlers:
                return "Dennis S."
            elif Casehandler in Rosa_Casehandlers:
                return "Rosa A."
            else:
                return "ZZ No Intake Para Assigned"

        df['Intake Paralegal'] = df.apply(lambda x: IntakeAssign(x['Primary Advocate']),axis = 1)

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
        "Intake Paralegal"
        ]]      
        
        #Preparing Excel Document
        
        #Split into different tabs
        allgood_dictionary = dict(tuple(df.groupby('Intake Paralegal')))
        
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

                ws.conditional_format('F2:K100000',{'type': 'blanks',
                                                 'format': problem_format})
                ws.conditional_format('G2:G100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Hold For Review',
                                                 'format': problem_format})
                ws.conditional_format('F2:F100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'No',
                                                 'format': problem_format})
                ws.conditional_format('L2:M100000',{'type': 'text',
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
