from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, EmploymentToolBox, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/IOIempQuarterly", methods=['GET', 'POST'])
def upload_IOIempQuarterly():
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
        
        last7 = df['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = df['Matter/Case ID#']
        df['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        move=df['Temp Hyperlinked Case #']
        df.insert(0,'Hyperlinked Case #', move)           
        del df['Temp Hyperlinked Case #']
        
        df.fillna('',inplace=True)
        
        
        #shorten branch names
        df['Assigned Branch/CC'] = df.apply(lambda x: DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis = 1)
        
        
        #Putting Employment Work in HRA Baskets
        
        df['HRA_Case_Coding'] = df.apply(lambda x: EmploymentToolBox.HRA_Case_Coding(x['Level of Service'],x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA IOI Employment Law Retainer?'],x['Matter/Case ID#']), axis = 1)        

        #Does case need special legal problem code?
                
        df['Special Legal Problem Code'] = df.apply(lambda x: EmploymentToolBox.SPLC_problem(x['Level of Service'],x['Special Legal Problem Code'],x['HRA_Case_Coding']), axis = 1)
        
        
        #Can case be reported based on income?

        df['Exclude due to Income?'] = df.apply(lambda x: EmploymentToolBox.Income_Exclude(x['Percentage of Poverty'],x['HRA IOI Employment Law If Client is Over 200% of FPL, Did you seek a waiver from HRA?']), axis=1)
        
        
        #DateMaker for Date Opened
        
        df['Open Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']),axis = 1)

        
        #DHCI form needed?
        
        
        df['Needs DHCI?'] = df.apply(lambda x: EmploymentToolBox.DHCI_Needed(x['HRA IOI Employment Law DHCI Form?'],x['Level of Service'],x['Open Construct']), axis=1)
        
         
        #Manipulable Dates (this seems like a mess, i would like to fix it later - Jay)            
        
        def Eligibility_Date(Effective_Date,Date_Opened):
            if Effective_Date != '':
                return Effective_Date
            else:
                return Date_Opened
        df['Eligibility_Date'] = df.apply(lambda x : Eligibility_Date(x['IOI HRA Effective Date (optional) (IOI 2)'],x['Date Opened']), axis = 1)
        
        df['Open Month'] = df['Eligibility_Date'].apply(lambda x: str(x)[:2])
        df['Open Day'] = df['Eligibility_Date'].apply(lambda x: str(x)[3:5])
        df['Open Year'] = df['Eligibility_Date'].apply(lambda x: str(x)[6:])
        df['Open Construct'] = df['Open Year'] + df['Open Month'] + df['Open Day']
        
        #DateMaker Substantial Activity FY21
        df['Subs Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HRA IOI Employment Law HRA Date Substantial Activity Performed 2021']),axis = 1)
                
        #Substantial Activity for Rollover FY21?

        df['Needs Substantial Activity?'] = df.apply(lambda x: EmploymentToolBox.Needs_Rollover(x['Open Construct'],x['HRA IOI Employment Law HRA Substantial Activity 2021'],x['Subs Construct'],x['Matter/Case ID#']), axis=1)
        
        #Unit of Service Calculator
        df['Units of Service'] = df.apply(lambda x: EmploymentToolBox.UoSCalculator(x['HRA_Case_Coding']),axis=1)
        
        #Reportable?
        
        df['Reportable?'] = df.apply(lambda x: EmploymentToolBox.ReportableTester(x['Exclude due to Income?'],x['Needs DHCI?'],x['Needs Substantial Activity?'],x['HRA_Case_Coding']),axis=1)
        
        #Assign Outcomes
        def AdviceOutcomeDate(HRAOutcome,HRAOutcomeDate,DateClosed):
            if HRAOutcomeDate == '' and HRAOutcome == 'Advice Given':
                return DateClosed
            else:
                return HRAOutcomeDate
                
        df['HRA IOI Employment Law HRA Outcome Date:'] = df.apply(lambda x: AdviceOutcomeDate(x['HRA IOI Employment Law HRA Outcome:'],x['HRA IOI Employment Law HRA Outcome Date:'],x['Date Closed']), axis = 1)
        
        def AdviceOutcome(HRAOutcome,Employment_Tier,CaseDisposition,HRAOutcomeDate):
            if HRAOutcome == '' and CaseDisposition== 'Closed' and Employment_Tier == 'Advice-No Retainer':
                return 'Advice Given'
            elif HRAOutcome == '' and CaseDisposition == 'Closed':
                return '**Needs Outcome**'
            elif HRAOutcome != '' and CaseDisposition == 'Closed'and HRAOutcomeDate == '':
                return '**Needs Outcome Date**'
            else:
                return HRAOutcome

        df['HRA Outcome'] = df.apply(lambda x: AdviceOutcome(x['HRA IOI Employment Law HRA Outcome:'], x['HRA IOI Employment Law IOI Employment Tier Category:'], x['Case Disposition'],x['HRA IOI Employment Law HRA Outcome Date:']), axis = 1)


        #Better names & HRA Names

        df['Employment Tier Category'] = df['HRA IOI Employment Law IOI Employment Tier Category:']
        
        df['Client Name'] = df['Full Person/Group Name (Last First)']
        
        df['Office'] = df['Assigned Branch/CC']
        
        df['Unique_ID'] = 'LSNYC'+df['Matter/Case ID#']
        
        df['Last_Initial'] = df['Client Last Name'].str[1]
        df['First_Initial'] = df['Client First Name'].str[1]
        
        df['Year_of_Birth'] = df['Date of Birth'].str[-4:]
        
        #gender
        def HRAGender (gender):
            if gender == 'Male' or gender == 'Female':
                return gender
            else:
                return 'Other'
        df['Gender'] = df.apply(lambda x: HRAGender(x['Gender']), axis=1)
        
        df['Country of Origin'] = ''
        
        #county=borough
        df['Borough'] = df['County of Residence']
        
        #household size etc.
        df['Household_Size'] = df['Number of People under 18'].astype(int) + df['Number of People 18 and Over'].astype(int)
        df['Number_of_Children'] = df['Number of People under 18']
        
        #Income Eligible?
        df['Annual_Income'] = df['Total Annual Income ']
        def HRAIncElig (PercentOfPoverty):
            if PercentOfPoverty > 200:
                return 'NO'
            else:
                return 'YES'
        df['Income_Eligible'] = df.apply(lambda x: HRAIncElig(x['Percentage of Poverty']), axis=1)
        
        def IncWaiver (eligible,waiverdate):
            if eligible == 'NO' and waiverdate != '':
                return 'Income'
            else:
                return ''
        df['Waiver_Type'] = df.apply(lambda x: IncWaiver(x['Income_Eligible'],x['HRA IOI Employment Law If Client is Over 200% of FPL, Did you seek a waiver from HRA?']), axis=1)
        
        def IncWaiverDate (waivertype,date):
            if waivertype != '':
                return date
            else:  
                return ''
                
        df['Waiver_Approval_Date'] = df.apply(lambda x: IncWaiverDate(x['Waiver_Type'],x['HRA IOI Employment Law Income Waiver Date']), axis = 1)
        

        #Other Cleanup
        
        
        df['Referral_Source'] = 'None'
        
        def ProBonoCase (branch, pai):
            if branch == "LSU" or pai == "Yes":
                return "YES"
            else:
                return "NO"
                
        def PriorEnrollment (casenumber):
            if casenumber in EmploymentToolBox.ReportedFY20:
                return 'FY 20'
            elif casenumber in EmploymentToolBox.ReportedFY19:
                return 'FY 19'
                
        df['Prior_Enrollment_FY'] = df.apply(lambda x:PriorEnrollment(x['Matter/Case ID#']), axis = 1)
                
        df['Pro_Bono'] = df.apply(lambda x:ProBonoCase(x['Assigned Branch/CC'], x['PAI Case?']), axis = 1)
       
        df['Service_Type_Code'] = df['HRA_Case_Coding'].apply(lambda x: x[:2] if x != 'Hold For Review' else '')

        df['Proceeding_Type_Code'] = df['HRA_Case_Coding'].apply(lambda x: x[3:] if x != 'Hold For Review' else '')
        
        df['Outcome'] = df['HRA IOI Employment Law HRA Outcome:']
        df['Outcome_Date'] = df['HRA IOI Employment Law HRA Outcome Date:']
        
        df['Seized_at_Border'] = ''
        
        df['Group'] = ''
        
        
        
        #sorting by borough and advocate
        df = df.sort_values(by=['Office','Primary Advocate'])
        
        
        
        #REPORTING VERSION Put everything in the right order
        df = df[['Unique_ID','Last_Initial','First_Initial','Year_of_Birth','Gender','Country of Origin','Borough','Zip Code','Language','Household_Size','Number_of_Children','Annual_Income','Income_Eligible','Waiver_Type','Waiver_Approval_Date','Eligibility_Date','Referral_Source','Service_Type_Code','Proceeding_Type_Code','Outcome','Outcome_Date','Seized_at_Border','Group','Prior_Enrollment_FY','Pro_Bono','Hyperlinked Case #','Office','Primary Advocate','Client Name','Level of Service','Legal Problem Code','Special Legal Problem Code','HRA_Case_Coding','Exclude due to Income?','Needs DHCI?','Needs Substantial Activity?','Units of Service','Reportable?']]
        

        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        problem_format = workbook.add_format({'bg_color':'yellow'})
        
        
        
        worksheet.set_column('C:BL',30)
        worksheet.conditional_format('E1:E100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
        worksheet.conditional_format('F1:F100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"***Needs SPLC***"',
                                                 'format': problem_format})
        worksheet.conditional_format('G1:G100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Income Waiver"',
                                                 'format': problem_format})
        worksheet.conditional_format('H1:H100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs DHCI"',
                                                 'format': problem_format})
        worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Substantial Activity in FY20"',
                                                 'format': problem_format})
        worksheet.conditional_format('J1:J100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
        worksheet.conditional_format('K1:K100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"**Needs Outcome**"',
                                                 'format': problem_format})
        worksheet.conditional_format('K1:K100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"**Needs Outcome Date**"',
                                                 'format': problem_format})
        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Formatted " + f.filename)

    return '''
    <!doctype html>
    <title>IOI Employment Quarterly</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Quarterly Formatter for IOI Employment Cases:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=IOI-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2020" target="_blank">"Grants Management IOI Employment (3474) Report"</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for IOI.</li> 
    <li>Once you have identified this file, click ‘IOI-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
