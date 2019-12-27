from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/UAHPLP", methods=['GET', 'POST'])
def upload_UAHPLP():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        test = pd.read_excel(f)
        
        test.fillna('',inplace=True)
        
        #Cleaning
        if test.iloc[0][0] == '':
            data_xls = pd.read_excel(f,skiprows=2)
        else:
            data_xls = pd.read_excel(f)
        
        data_xls.fillna('',inplace=True)
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = data_xls['Matter/Case ID#']
        data_xls['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del data_xls['Matter/Case ID#']
        move=data_xls['Temp Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Bronx Legal Services','BxLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Brooklyn Legal Services','BkLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Queens Legal Services','QLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Manhattan Legal Services','MLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Staten Island Legal Services','SILS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Legal Support Unit','LSU')
        
                
        #Manipulable Dates               
        
        #data_xls['Open Month'] = data_xls['Date Opened'].apply(lambda x: str(x)[:2])
        #data_xls['Open Day'] = data_xls['Date Opened'].apply(lambda x: str(x)[3:5])
        #data_xls['Open Year'] = data_xls['Date Opened'].apply(lambda x: str(x)[6:])
        #data_xls['Open Construct'] = data_xls['Open Year'] + data_xls['Open Month'] + data_xls['Open Day']
        
        data_xls['Elig Month'] = data_xls['HAL Eligibility Date'].apply(lambda x: str(x)[:2])
        data_xls['Elig Day'] = data_xls['HAL Eligibility Date'].apply(lambda x: str(x)[3:5])
        data_xls['Elig Year'] = data_xls['HAL Eligibility Date'].apply(lambda x: str(x)[6:])
        data_xls['Elig Construct'] = data_xls['Elig Year'] + data_xls['Elig Month'] + data_xls['Elig Day']
        
        data_xls['Close Month'] = data_xls['Date Closed'].apply(lambda x: str(x)[:2])
        data_xls['Close Day'] = data_xls['Date Closed'].apply(lambda x: str(x)[3:5])
        data_xls['Close Year'] = data_xls['Date Closed'].apply(lambda x: str(x)[6:])
        data_xls['Close Construct'] = data_xls['Close Year'] + data_xls['Close Month'] + data_xls['Close Day']
        
        #opened during current fiscal year with no HRA eligibility datetime
        
        def EligTest (EligCons):
            if EligCons == '':
                EligConst = 0
                return 'Needs Eligibility Date'
            EligCons = int(EligCons)
            if EligCons < 20190701 or EligCons > 20200630:
                return 'Eligibility Date Out of Contract Year'
            elif EligCons >= 20190701 and EligCons <= 20200630:
                return 'All Good!'
            else:
                return 'Something is Weird'
                
        data_xls['Eligibility Tester'] = data_xls.apply(lambda x: EligTest(x['Elig Construct']), axis=1)
        
                       
        #no type of case/proceeding info entered
        
        def NoCaseType (CaseType):
            if CaseType != '':
                return 'All Good!'
            elif CaseType == '':
                return 'Needs Housing Case Type'
            else:
                return 'Something is Weird'
        
        data_xls['Blank Case Type Tester'] = data_xls.apply(lambda x: NoCaseType(x['Housing Type Of Case']),axis=1)
        
        # no hra consent form (blank or checked no)
        
        def NoConsent (Consent):
            if Consent != ' ' and Consent != 'No':
                return 'All Good!'
            elif Consent == ' ' or Consent == 'No':
                return 'Needs HRA Release/Consent Form'
            else:
                return 'Something is Weird'
        
        data_xls['HRA Release/Consent Tester'] = data_xls.apply(lambda x: NoConsent(x['HRA Release?']),axis=1)
        
        #no  PA case number for which DHCI is also blank or checked no
        
        def DHCIorPub (DHCI,PubAssist,VerifType):
            if DHCI == 'Yes':
                return 'All Good!'
            elif PubAssist != '':
                return 'All Good!'
            elif VerifType == 'DHCI Form' or VerifType == 'Active CA/SNAP':
                return 'All Good!'
            elif PubAssist == '' and DHCI == 'No':
                return 'Needs DHCI'
            elif PubAssist == '' and DHCI == ' ':
                return 'Needs DHCI'
            
            else:
                return 'Something is Weird'
        
        data_xls['Income Verification Tester'] = data_xls.apply(lambda x: DHCIorPub(x['Housing Signed DHCI Form'],x['Gen Pub Assist Case Number'],x['Housing Income Verification']),axis=1)
            
        
        #no level of service or is in hold for review
        
        def NoLoS (LoS):
            if LoS != '' and LoS != 'Hold For Review':
                return 'All Good!'
            elif LoS == '' or LoS == 'Hold For Review':
                return 'Needs Level of Service'
            else:
                return 'Something is Weird'
        
        data_xls['Level of Service Tester'] = data_xls.apply(lambda x: NoLoS(x['Housing Level of Service']),axis=1)
        
        ##eligibility date in current fiscal year and closed during current fiscal year with no HRA/Housing View Outcome
        
        def NoOutcomeOrDate (CloseCons,Outcome,OutcomeDate):
            if CloseCons == '':
                CloseCons = 0
            CloseCons = int(CloseCons)
            if Outcome != '' and OutcomeDate != '':
                return 'All Good!'
            elif CloseCons == 0:
                return 'Case Not Closed'
            elif CloseCons < 20190701 or CloseCons > 20200630:
                return 'Close Date Outside of FY20'
            elif Outcome == '' and OutcomeDate == '':
                return 'Needs Outcome & Outcome Date'
            elif Outcome == '':
                return 'Needs Outcome'
            elif OutcomeDate == '':
                return 'Needs Outcome Date'
            else:
                return 'Something is Weird'
        
        data_xls['Outcome Tester'] = data_xls.apply(lambda x: NoOutcomeOrDate(x['Close Construct'],x['Housing Outcome'],x['Housing Outcome Date']),axis=1)
        

        #Is everything okay with a case? (delete if so?)
        
        def TesterTester (EligTester,CaseTypeTester,ReleaseTester,IncomeTester,LoSTester,OutcomeTester):
            if EligTester == 'All Good!' and CaseTypeTester == 'All Good!' and ReleaseTester == 'All Good!' and IncomeTester == 'All Good!' and LoSTester == 'All Good!' and OutcomeTester == 'Case Not Closed': 
                return "All Good!!!"
            elif EligTester == 'All Good!' and CaseTypeTester == 'All Good!' and ReleaseTester == 'All Good!' and IncomeTester == 'All Good!' and LoSTester == 'All Good!' and OutcomeTester == 'All Good!': 
                return "All Good!!!"    
            else:
                return 'Case Needs Attention'
                
        data_xls['Tester Tester'] = data_xls.apply(lambda x: TesterTester(x['Eligibility Tester'],x['Blank Case Type Tester'],x['HRA Release/Consent Tester'],x['Income Verification Tester'],x['Level of Service Tester'],x['Outcome Tester']),axis=1)
        
        data_xls = data_xls[data_xls['Tester Tester'] != 'All Good!!!']
        
        #sort by assigned branch and case handler
        
        data_xls = data_xls.sort_values(by=['Primary Advocate'])
        data_xls = data_xls.sort_values(by=['Assigned Branch/CC'])
        
        #Put everything in the right order
        
        data_xls = data_xls[['Hyperlinked Case #','Assigned Branch/CC','Primary Advocate','Tester Tester','Eligibility Tester','HAL Eligibility Date','Blank Case Type Tester','Housing Type Of Case','HRA Release/Consent Tester','HRA Release?','Income Verification Tester','Gen Pub Assist Case Number','Housing Signed DHCI Form','Housing Income Verification','Level of Service Tester','Housing Level of Service','Outcome Tester','Housing Outcome','Housing Outcome Date',""""""'Client First Name','Client Last Name','Date Opened','Date Closed','Case Disposition','Street Address','Apt#/Suite#','City','State','Zip Code','Referral Source','Gen Case Index Number','Housing Years Living In Apartment','Close Reason','Primary Funding Code','Group','Housing Building Case?','Secondary Funding Codes','Legal Problem Code','Housing Posture of Case on Eligibility Date','Housing Tenant’s Share Of Rent','Housing Total Monthly Rent','Total Time For Case','Outcome','Date of Birth','Social Security #','Housing Number Of Units In Building','Housing Form Of Regulation','Number of People 18 and Over','Number of People under 18','Percentage of Poverty','Housing Date Of Waiver Approval','Housing TRC HRA Waiver Categories','Housing HPLP Household Category','Housing Subsidy Type','Language','Total Annual Income ','Housing Proof Public Assistance','Housing Verification Of Income','Housing Funding Note','Caseworker Name','Housing Activity Indicators','Housing Services Rendered to Client','Income Types','Service Date','Housing Income Verification']]      
        
        #Preparing Excel Document
        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Sheet1',index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        regular_format = workbook.add_format({'font_color':'black'})
        problem_format = workbook.add_format({'bg_color':'yellow'})
        
        
        worksheet.set_column('A:A',20,link_format)
        worksheet.set_column('B:B',19)
        worksheet.set_column('C:BL',30)
        worksheet.conditional_format('E1:E100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Eligibility Date"',
                                                 'format': problem_format})
        worksheet.conditional_format('E1:E100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Eligibility Date Out of Contract Year"',
                                                 'format': problem_format})
        worksheet.conditional_format('G1:G100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Housing Case Type"',
                                                 'format': problem_format})
        worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs HRA Release/Consent Form"',
                                                 'format': problem_format})
        worksheet.conditional_format('K1:K100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs DHCI"',
                                                 'format': problem_format})                                        
        worksheet.conditional_format('O1:O100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Level of Service"',
                                                 'format': problem_format})                                        
        worksheet.conditional_format('Q1:Q100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Outcome & Outcome Date"',
                                                 'format': problem_format})                                     
        worksheet.conditional_format('Q1:Q100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Outcome"',
                                                 'format': problem_format})
        worksheet.conditional_format('Q1:Q100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Outcome Date"',
                                                 'format': problem_format})
        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>UAHPLP</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Prepare a UA/HPLP Cleanup Report:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1963" target="_blank">HPLP/UAC Internal Report</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for UA/HLPL cleanup.</li> 
    <li>Once you have identified this file, click ‘Clean!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
