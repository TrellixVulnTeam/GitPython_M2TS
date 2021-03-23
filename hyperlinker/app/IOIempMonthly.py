from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, EmploymentToolBox, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd



@app.route("/IOIempMonthly", methods=['GET', 'POST'])
def upload_IOIempMonthly():
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
        
         
        #Manipulable Dates               
        
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

        df['Client Name'] = df['Full Person/Group Name (Last First)']
        
        df['Office'] = df['Assigned Branch/CC']
        

        
        df['Outcome'] = df['HRA IOI Employment Law HRA Outcome:']
        df['Outcome_Date'] = df['HRA IOI Employment Law HRA Outcome Date:']
        
        
        #Cleaning Output
        
        df = df[['Hyperlinked Case #','Office','Primary Advocate','Client Name','Level of Service','Special Legal Problem Code','Exclude due to Income?','Needs DHCI?','Needs Substantial Activity?','Language','HRA Outcome','HRA_Case_Coding','Units of Service','Reportable?']]
        
        #sorting by borough and advocate
        df = df.sort_values(by=['Office','Primary Advocate'])
        
        borough_dictionary = dict(tuple(df.groupby('Office')))
           
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                worksheet = writer.sheets[i]
                worksheet.set_column('A:A',20,link_format)
                worksheet.set_column('B:B',19)
                worksheet.set_column('C:BL',30)
                worksheet.freeze_panes(1,1)
                
                
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
                                                 'value': '"Needs Substantial Activity in FY21"',
                                                 'format': problem_format})
                worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Substantial Activity Date"',
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
        
        output_filename = f.filename
        
        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)
           

        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleanup " + f.filename)

    return '''
    <!doctype html>
    <title>IOI Employment Monthly</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Monthly Cleanup for IOI Employment Cases:</h1>
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