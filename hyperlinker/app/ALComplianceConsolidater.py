#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os

@app.route("/ALComplianceConsolidater", methods=['GET', 'POST'])
def ALComplianceConsolidater():
    #upload file from computer
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        #turn the excel file into a dataframe, but skip the top 2 rows if they are blank
        test = pd.read_excel(f)
        test.fillna('',inplace=True)
        if test.iloc[0][0] == '':
            df = pd.read_excel(f,skiprows=2)
        else:
            df = pd.read_excel(f)

        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Checkbox determines if it does lsu or all non-LSU***for editing
        '''if request.form.get('Caseworker'):
            df = df[df['Assigned Branch/CC'] == "Legal Support Unit"]
        elif request.form.get('QLS'):
            df = df[df['Assigned Branch/CC'] == "Queens Legal Services"]
        elif request.form.get('MLS'):
        else:
            df = df'''
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)     
        
        #Remove duplicate Case ID #s
        
        #df = df.drop_duplicates(subset='Hyperlinked Case #', keep = 'first')
        
        #This is where all the functions happen:
               
        
        #200 percent of poverty income eligible
        #if percentage of poverty is > 200 AND
        #LSC or CSR = yes***removal unclear
        
        def TwoHundredPercentTester(PovPercent,LSCEligible,CSREligible,ComplianceCheck):
            if PovPercent > 200 and LSCEligible == 'Yes' and ComplianceCheck != "Yes":
                return "Needs Review"
            elif PovPercent >200 and CSREligible == 'Yes' and ComplianceCheck != "Yes":
                return "Needs Review"
            else:
                return ''
        
        df['200% of Poverty Tester'] = df.apply(lambda x : TwoHundredPercentTester(x['Percentage of Poverty'],x['LSC Eligible?'],x['CSR Eligible'],x['Compliance Check 200 Poverty Income Eligible']),axis=1)
        
        
        #125-200 poverty income eligible
        #percentage of poverty is between 125 and 200 AND 
        #Income eligible = no ***keep
        
        def OneTwentyFivePercentTester(PovPercent,IncomeEligible, ComplianceCheck):
            if PovPercent > 125 and PovPercent < 200 and IncomeEligible == 'No' and ComplianceCheck != "Yes":
                return "Needs Review"
            else:
                return ''
        
        df['125-200% of Poverty Tester'] = df.apply(lambda x : OneTwentyFivePercentTester(x['Percentage of Poverty'],x['Income Eligible'],x['Compliance Check 125 to 200 Poverty Income Ineligible']),axis=1)
        
        
        '''
        #No Age for Client
        #Not a Group
        #DOB Status = Refused or Unknown (approximate or blank are fine)***keep
        
        def NoAgeTester (DOBStatus,Group):
            if DOBStatus == "Unknown" and Group != "Yes":
                return "Needs Review"
            else:
                return ""
        
        df['No Age for Client Tester'] = df.apply(lambda x : NoAgeTester(x['DOB Information'],x['Group']),axis=1)
        '''
        #Citizenship and Immigration Compliance
        #Is case closed?
        #is attestation on file = no or unanswered
        #and staff verified non-citizenship documentation = no or unanswered
        #AND either
        #did staff meet client in person = yes or unanswered
        #or close reason starts with = F,G,H,I,or L***keep
        
        
        def CitImmALTester(CitizenshipStatus, ImmigrationStatus):
            if CitizenshipStatus == "Non-Citizen":
                if ImmigrationStatus != "Asylee" and ImmigrationStatus != 'Lawful Permanent Resident (LPR)' and ImmigrationStatus != 'Refugee':
                
                    return 'Needs Review'
            else:
                return ''
        
        
        #***anyone that's not a citizen, not asylee, not refugee, not greencard holder
        #"anti-abuse notes" include in another column
        
        #***sort by category of issue instead of caseworker like MLS compliance
            
        df['Citizenship & Immigration Tester'] = df.apply(lambda x : CitImmALTester(x['Citizenship Status'],x['Immigration Status']),axis=1)
        
        #Intake Splitter:
        
        IntakeList = [
        "Ikram, Nabah",
        "Suriel, Sal",
        "Rodney, Gabby A",
        "Paz, Alexis",
        "Alexis, Jennifer",
        "Yeh, Victoria",
        "Navas, Keyra M",
        "Dong, Sean",
        "Baldova, Maria",
        "Djourab, Atteib",
        "Ortega, Luis",
        "Pierre, Haenley",
        "Espinal, Wendy",
        "Cordero, Natalie D",
        "Agarwala, Shelly",
        
        "Matero, Dante C",
        "Espinal, Wendy"]
        
        #make a tester tester (for year end)
        
          
        def TesterTester(TwoHundredPercentTester,OneTwentyFivePercentTester,CitImmTester):
            if TwoHundredPercentTester != "" or OneTwentyFivePercentTester != "" or CitImmTester != "":
                return "Needs Review"
            else:
                return ""
        df['Tester Tester'] = df.apply(lambda x : TesterTester(x['200% of Poverty Tester'],x['125-200% of Poverty Tester'],x['Citizenship & Immigration Tester']),axis=1)        
        
        
        #change name of column for intake worker
        df['Intake Caseworker'] = df['Caseworker Name']
        
        df['LSNYC']= "LSNYC"
        
        def IntakeFilter (IntakePerson):
            if IntakePerson in IntakeList:
                return "On hotline"
            else:
                return "Not on hotline"
                
        df['On Hotline?'] = df.apply(lambda x :IntakeFilter(x['Intake Caseworker']),axis=1) 
        
        
        
        #Putting everything in the right order
        df = df[df['Tester Tester'] != ""]    
        df = df[df['On Hotline?'] != 'Not on hotline']
        
        df = df[['Hyperlinked CaseID#',
        'Assigned Branch/CC',
        'Primary Advocate Name',
        'Client First Name',
        'Client Last Name',
        #'Tester Tester',
        '200% of Poverty Tester',
        '125-200% of Poverty Tester',
        #'No Age for Client Tester',
        'Citizenship & Immigration Tester',
        #'Date Opened',
        #'Case Status',
        'Intake Caseworker',
        'Citizenship Status',
        'Immigration Status',
        'Anti-Abuse Statutes Documentation Note (Notes)',
        'Intake - Did You Provide Legal Advice?',
        'Financial Eligibility Override Reason',
        'Financial Override Notes (Notes)',
        #'Compliance Check 125 to 200 Poverty Income Ineligible',
        #'Compliance Check 200 Poverty Income Eligible',
        #'Compliance Check Citizenship and Immigration',
        #
        #'Attestation on File?',
        #'Staff Verified Non-Citizenship Documentation',
        #'Did any Staff Meet Client in Person?',
        #'Close Reason',
        #'Date Closed',
        #'Percentage of Poverty',
        #'Income Eligible',
        #'DOB Information',
        #'Group'
        ]]
    
        #Choose which data to return
        if request.form.get('Aggregate Data'):
            
            TabSplitValue = df['LSNYC']
            print("Aggregate Data = on")
        else:
            
            
            
            TwoHundredPoverty_df = df[df['200% of Poverty Tester'] == "Needs Review"]
            OneTwentyFivePovertydf = df[df['125-200% of Poverty Tester'] == "Needs Review"]
            Citizenship_df = df[df['Citizenship & Immigration Tester'] == "Needs Review"]
            Override_df = df[df['Financial Eligibility Override Reason'] != ""]
            
            Override_df = Override_df[['Hyperlinked CaseID#',
            'Assigned Branch/CC',
            'Primary Advocate Name',
            'Client First Name',
            'Client Last Name',
            'Financial Eligibility Override Reason',
            'Financial Override Notes (Notes)',
            'Intake Caseworker',
            ]]
                
            problem_type_dictionary = {'200% Poverty': TwoHundredPoverty_df , '125% Poverty': OneTwentyFivePovertydf , 'Citizenship': Citizenship_df , 'Override Reason': Override_df}
            
            
            
            TabSplitValue = df['Intake Caseworker']
            print("Split by Type of Problem")
                    

        df = df.sort_values(by=['Intake Caseworker'])
        

        #Preparing Excel Document
        #bounce worksheets back to excel
        def save_xls(TabSplit, path):
        
            #global output_filename
            #output_filename = f.filename     
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            
            #Create dictionary of tabs, each named after each unique value of 'Intake Case Worker'
            tab_dict_df = problem_type_dictionary
            
            #creates for loop revolving around the split value, ICW or LSNYC(no split)
            
            for splitval in tab_dict_df:
            
                #writes document to excel
                tab_dict_df[splitval].to_excel(writer, splitval, index = False)
                
                #creates ability to format tabs
                worksheet = writer.sheets[splitval]
                workbook = writer.book
                worksheet.freeze_panes(1, 2) 
                
                #create format that will make case #s look like links
                link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                worksheet.autofilter('A1:Z1')
                B2KFullRange='B2:K'+str(tab_dict_df[splitval].shape[0]+1) 
                print("BAHRowRange is "+ str(B2KFullRange)) 
                
                for column in df:
                    column_width = max(df[column].astype(str).map(len).max(), len(column))
                    #print("I am column width" + str(column_width))
                    col_idx = df.columns.get_loc(column)
                    #print("I am col idx" + str(col_idx))
                    writer.sheets[splitval].set_column(col_idx, col_idx, column_width)
                
                #assign new format to column A
                worksheet.set_column('A:A',20,link_format)
                worksheet.set_column('L:L',50)
                worksheet.conditional_format(B2KFullRange,{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
                worksheet.conditional_format('B1:Z1',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Tester',
                                                 'format': problem_format})                                                 
            writer.save()
        
        save_xls(TabSplit = TabSplitValue, path = "app\\sheets\\" + f.filename)
        return send_from_directory('sheets',f.filename, as_attachment = True, attachment_filename = "Access Line " + f.filename) 
        
        '''#send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Hyperlinked " + f.filename)'''
        
        #***#
       
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Access Line Compliance Consolidater</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Summary of Access Line Compliance Reports</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Comply!!>
    
    </br>
        
    <input type="checkbox" id="Aggregate Data" name="Aggregate Data" value="Aggregate Data">
    <label for="Aggregate Data"> Aggregate Data</label><br>
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2305" target="_blank">Compliance Consolidated Report</a>.</li>
    <li>This tool defaults to plitting by Intake Caseworker  last updated 05/2022.</li>
    <li>The Aggregate data button will not split the data.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
