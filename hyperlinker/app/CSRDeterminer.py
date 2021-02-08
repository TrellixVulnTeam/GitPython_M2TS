#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os

@app.route("/CSRDeterminer", methods=['GET', 'POST'])
def CSRDeterminer():
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
        

        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)     
        
        
        #No Legal Assistance Documented Compliance Check

        def NoAssistanceTester(AssistanceDocumented):
            if AssistanceDocumented == "Yes":
                return ""
            else:
                return 'CSR No'
                 
        df['No Legal Assistance Documented Tester'] = df.apply(lambda x : NoAssistanceTester(x['CSR: Is Legal Assistance Documented?']),axis=1)
        
        
        #Untimely Closed
        
        def UntimelyClosedTester(TimelyClosing):
            if TimelyClosing == "Yes":
                return ""
            else:
                return 'CSR No'
                 
        df['Untimely Closed Tester'] = df.apply(lambda x : UntimelyClosedTester(x['CSR: Timely Closing?']),axis=1)
        
        #LSC Compliant
        
        def LSCTester(LSC):
            if LSC == "Yes":
                return ""
            else:
                return 'CSR No'

        df['LSC Tester'] = df.apply(lambda x : LSCTester(x['LSC Eligible?']),axis=1)
        
        #Citizenship and Immigration Compliance
        #is attestation on file = no or unanswered
        #and staff verified non-citizenship documentation = no or unanswered
        #AND either
        #did staff meet client in person = yes or unanswered
        #or close reason starts with = F,G,H,I,or L
        
        
        def CitImmTester(AttestationOnFile,StaffVerified,ClientInPerson,CloseReason):
            CloseReason = str(CloseReason)   
            if AttestationOnFile == "Yes" or StaffVerified == "Yes":
                return ''
            elif CloseReason.startswith(('A','B')) == True:
                return ''
            else:
                return 'CSR No'

        df['Citizenship & Immigration Tester'] = df.apply(lambda x : CitImmTester(x['Attestation on File?'],x['Staff Verified Non-Citizenship Documentation'],x['Did any Staff Meet Client in Person?'],x['Close Reason']),axis=1)
        
        def CSRTester (LSCTester,LegalAssistTester,UntimelyTester,CitImmTester):
            if LSCTester == 'CSR No' or LegalAssistTester == 'CSR No' or UntimelyTester == 'CSR No' or CitImmTester == 'CSR No':
                return "No"
            else:
                return "Yes"
        
        df['Python CSR Tester'] = df.apply(lambda x : CSRTester(x['LSC Tester'],x['No Legal Assistance Documented Tester'],x['Untimely Closed Tester'],x['Citizenship & Immigration Tester']),axis=1)
        
        def AgreementTester (LegalServerCSR,PythonCSR):
            if LegalServerCSR == 'Yes' and PythonCSR == 'Yes':
                return ''
            elif LegalServerCSR == 'No' and PythonCSR == 'No':
                return ''
            elif LegalServerCSR == 'Yes' and PythonCSR == 'No':
                return 'LSYesPythonNO'
            elif LegalServerCSR == 'No' and PythonCSR == 'Yes':
                return 'LSNoPythonYes'
            else:
                return 'How?!'
                
        df['Agreement Tester'] = df.apply(lambda x : AgreementTester(x['CSR Eligible'],x['Python CSR Tester']),axis=1)
        
        #Putting everything in the right order
        
        df['Case?'] = "Yes"
        
        df = df[['Hyperlinked CaseID#','Assigned Branch/CC','Primary Advocate Name','LSC Tester','No Legal Assistance Documented Tester','Untimely Closed Tester','Citizenship & Immigration Tester','Python CSR Tester','CSR Eligible','Agreement Tester','Staff Verified Non-Citizenship Documentation','Did any Staff Meet Client in Person?','Close Reason','Date Closed','Date Opened','CSR: Is Legal Assistance Documented?','LSC Eligible?','CSR Eligible','Income Eligible','CSR: Timely Closing?','Level of Service','Case?','Retainer on File','Was Timely Closed overridden?','Percentage of Poverty','Attestation on File?']]
        
        
        
        
        #Preparing Excel Document

        borough_dictionary = dict(tuple(df.groupby('Case?')))

        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                
                
                
                worksheet = writer.sheets[i]
                

                worksheet.set_column('A:A',15,link_format)
                worksheet.set_column('B:P',20)
                
                    
                worksheet.conditional_format('D2:BO100000',{'type': 'text',
                                                         'criteria': 'containing',
                                                         'value': 'CSR No',
                                                         'format': problem_format})
                worksheet.freeze_panes(1,1)
            writer.save()
        output_filename = f.filename

        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)

        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Python " + f.filename)
        
        #***#
       
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Determine CSR Status</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Determine CSR Status</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Determine!>
    
    </br>
    </br>

    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This should tell you whether a case is in fact eligible for the Case Services Report.</li>
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2305" target="_blank">Compliance Consolidated Report</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
