#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os

@app.route("/ComplianceConsolidater", methods=['GET', 'POST'])
def ComplianceConsolidater():
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
        
        #Remove duplicate Case ID #s
        
        #df = df.drop_duplicates(subset='Hyperlinked Case #', keep = 'first')
        
        #This is where all the functions happen:
        
        
        #No Legal Assistance Documented Compliance Check
        #Is Case Closed?
        #Is "CSR CSR: Is Legal Assistance Document? = No"
        
        def NoAssistanceTester(DateClosed,NoAssistance):
            if DateClosed != "nan" and NoAssistance == "No":
                return "Needs Review"
            else:
                return ''
                 
        df['No Legal Assistance Documented Tester'] = df.apply(lambda x : NoAssistanceTester(x['Date Closed'],x['CSR: Is Legal Assistance Documented?']),axis=1)
        
        #No Time Entered for 90 Days
        #just open cases
        #is 'Age in Days' larger than 90? (need to de-duplicate to get just cases with most recent service date?)
        
        def AgeTester(DateClosed,AgeInDays):
            DateClosed = str(DateClosed)
            if DateClosed == "" and AgeInDays > 90:
                return 'Needs Review'
            else:
                return ''
        
        df['No Time Entered for 90 Days Tester'] = df.apply(lambda x : AgeTester(x['Date Closed'],x['Age in Days']),axis=1)
        
        
        #200 percent of poverty income eligible
        #if percentage of poverty is > 200 AND
        #LSC or CSR = yes
        
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
        #Income eligible = no
        
        def OneTwentyFivePercentTester(PovPercent,IncomeEligible, ComplianceCheck):
            if PovPercent > 125 and PovPercent < 200 and IncomeEligible == 'No' and ComplianceCheck != "Yes":
                return "Needs Review"
            else:
                return ''
        
        df['125-200% of Poverty Tester'] = df.apply(lambda x : OneTwentyFivePercentTester(x['Percentage of Poverty'],x['Income Eligible'],x['Compliance Check 125 to 200 Poverty Income Ineligible']),axis=1)
        
        
        #LSC or CSR No on Funding Code 4000
        #Primary or secondary funding code = 4000 AND
        # CSR eligible OR LSC eligible = no
        
        def LSCCSRTester (PrimaryFundingCode,SecondaryFundingCode,LSCEligible,CSREligible):
            PrimaryFundingCode = str(PrimaryFundingCode)
            SecondaryFundingCode = str(SecondaryFundingCode)
            
            if PrimaryFundingCode.startswith('4000') == True and LSCEligible == "No":
                return "Needs Review"
            elif PrimaryFundingCode.startswith('4000') == True and CSREligible == "No":
                return "Needs Review"
            elif SecondaryFundingCode.startswith('4000') == True and LSCEligible == "No":
                return "Needs Review"
            elif SecondaryFundingCode.startswith('4000') == True and CSREligible == "No":
                return "Needs Review"
            else:
                return ''

        df['Funding Code 4000 Tester'] = df.apply(lambda x : LSCCSRTester(x['Primary Funding Codes'],x['Secondary Funding Codes'],x['LSC Eligible?'],x['CSR Eligible']),axis=1)
        
        #No Age for Client
        #Not a Group
        #DOB Status = Refused or Unknown (approximate or blank are fine)
        
        def NoAgeTester (DOBStatus,Group):
            if DOBStatus == "Unknown" and Group != "Yes":
                return "Needs Review"
            else:
                return ""
        
        df['No Age for Client Tester'] = df.apply(lambda x : NoAgeTester(x['DOB Information'],x['Group']),axis=1)
        
        #Untimely Closed
        #Closed Cases
        #CSR: Timely Closing = no
        
        def UntimelyClosedTester(DateClosed,TimelyClosing,ComplianceCheck):
            if DateClosed != "nan" and TimelyClosing == "No" and ComplianceCheck != "Yes":
                return "Needs Review"
            else:
                return ''
                 
        df['Untimely Closed Tester'] = df.apply(lambda x : UntimelyClosedTester(x['Date Closed'],x['CSR: Timely Closing?'],x['Compliance Check Untimely Closed']),axis=1)
        
        #Untimely Closed Overridden
        #is case closed?
        #CSR Timely Closing = YES AND
        #Was timely closed overridden = yes
        #and compliance check functionality
        
        def UntimelyClosedOverriddenTester(DateClosed,TimelyClosing,TimelyClosingOverridden,ComplianceCheck):
            if DateClosed != "nan" and TimelyClosing == "Yes" and TimelyClosingOverridden == "Yes" and ComplianceCheck != "Yes":
                return "Needs Review"
            else:
                return ''
                 
        df['Untimely Closed Overridden Tester'] = df.apply(lambda x : UntimelyClosedOverriddenTester(x['Date Closed'],x['CSR: Timely Closing?'],x['Was Timely Closed overridden?'],x['Compliance Check Untimely Closed Overwritten']),axis=1)

        
        #Citizenship and Immigration Compliance
        #Is case closed?
        #is attestation on file = no or unanswered
        #and staff verified non-citizenship documentation = no or unanswered
        #AND either
        #did staff meet client in person = yes or unanswered
        #or close reason starts with = F,G,H,I,or L
        
        
        def CitImmTester(AttestationOnFile,StaffVerified,ClientInPerson,CloseReason,ComplianceCheck):
            CloseReason = str(CloseReason)
            
            if AttestationOnFile == "Yes" or StaffVerified == "Yes":
                return ''
            elif ComplianceCheck == "Yes":
                return ''
            elif CloseReason.startswith(('A','B')) == True and ClientInPerson == 'No':
                return ''
            else:
                return 'Needs Review'
            
            
            
            
            
        df['Citizenship & Immigration Tester'] = df.apply(lambda x : CitImmTester(x['Attestation on File?'],x['Staff Verified Non-Citizenship Documentation'],x['Did any Staff Meet Client in Person?'],x['Close Reason'],x['Compliance Check Citizenship and Immigration']),axis=1)
        
        #Active Advocate Tester
        
        def ActiveAdvocateTester(ActiveAdvocate):
            if ActiveAdvocate != "Yes":
                return "Needs Review"
            else:
                return ''
                 
        df['Active Advocate Tester'] = df.apply(lambda x : ActiveAdvocateTester(x['Login Active']),axis=1)
        
        
        #also add in code related to checking the boxes they check if something has been compliance-reviewed
        
        #Make a tester tester
        def TesterTester(NoAssistanceTester,AgeTester,TwoHundredPercentTester,OneTwentyFivePercentTester,LSCCSRTester,NoAgeTester,UntimelyClosedTester,UntimelyClosedOverriddenTester,CitImmTester,ActAdvTester):
            if NoAssistanceTester != "" or AgeTester != "" or TwoHundredPercentTester != "" or OneTwentyFivePercentTester != "" or LSCCSRTester != "" or NoAgeTester != "" or UntimelyClosedTester != "" or UntimelyClosedOverriddenTester != "" or CitImmTester != "" or ActAdvTester != "":
                return "Needs Review"
            else:
                return ""
        df['Tester Tester'] = df.apply(lambda x : TesterTester(x['No Legal Assistance Documented Tester'],x['No Time Entered for 90 Days Tester'],x['200% of Poverty Tester'],x['125-200% of Poverty Tester'],x['Funding Code 4000 Tester'],x['No Age for Client Tester'],x['Untimely Closed Tester'],x['Untimely Closed Overridden Tester'],x['Citizenship & Immigration Tester'],x['Active Advocate Tester']),axis=1)        
        
        #remove cases that don't need compliance review
        df = df[df['Tester Tester'] == "Needs Review"]
        
        #Putting everything in the right order
        
        df = df[['Hyperlinked CaseID#','Assigned Branch/CC','Primary Advocate Name','Client First Name','Client Last Name','No Legal Assistance Documented Tester','No Time Entered for 90 Days Tester','200% of Poverty Tester','125-200% of Poverty Tester','Funding Code 4000 Tester','No Age for Client Tester','Untimely Closed Tester','Untimely Closed Overridden Tester','Citizenship & Immigration Tester','Active Advocate Tester','Caseworker Name','Compliance Check Untimely Closed','Compliance Check Untimely Closed Overwritten','Compliance Check 125 to 200 Poverty Income Ineligible','Compliance Check 200 Poverty Income Eligible','Compliance Check Citizenship and Immigration','Attestation on File?','Staff Verified Non-Citizenship Documentation','Did any Staff Meet Client in Person?','Close Reason','Date Closed','CSR: Is Legal Assistance Documented?','Age in Days','Percentage of Poverty','LSC Eligible?','CSR Eligible','Income Eligible','Primary Funding Codes','Secondary Funding Codes','DOB Information','Group','CSR: Timely Closing?','Was Timely Closed overridden?']]
        
        
        
        
        #Preparing Excel Document

        borough_dictionary = dict(tuple(df.groupby('Assigned Branch/CC')))

        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False, header = False,startrow=1)
                workbook = writer.book
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                header_format = workbook.add_format({
                'text_wrap':True,
                'bold':True,
                'valign': 'middle',
                'align': 'center'
                })
                
                
                worksheet = writer.sheets[i]
                
                #Add column header data back in
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                worksheet.autofilter('A1:O1')
                worksheet.set_column('A:A',15,link_format)
                worksheet.set_column('B:O',20)
                worksheet.set_column('P:ZZ',0)
                worksheet.set_row(0,60)
                worksheet.write_comment('F1',
                    'No Legal Assistance Documented: If we close a case as something other than ZZ, we must have some documentation of the legal assistance provided. For these, whoever closed the case may have misunderstood the meaning of “legal assistance” or “documented.” Please review and add the legal assistance documented in case notes. If there was no legal assistance, change the closing code to ZZ.',
                    {'height':70,'width':500,'x_offset':-30})
                worksheet.write_comment('G1',
                    'No Time in 90 Days: These are cases for which no time has been entered in 90 days. Please make sure these are all active. If they are active and pending in a court or other forum, enter a case note stating this. We have come across cases opened as long as four or five years ago for which no time had been entered since the year of opening, and which should have been closed in the same year they were opened.',
                    {'height':70,'width':500})
                worksheet.write_comment('H1',
                    '200% of Poverty Income Eligible: Please confirm that the overrides were completed properly in these cases; these are rarely correct, but it is possible. If they were properly done, there is no need to do anything else.',
                    {'height':70,'width':500})
                worksheet.write_comment('I1',
                    '125-200 Poverty Income Ineligible with Type of Income: These are cases where the client is between 125% and 200% of poverty and the financial override was not completed. In most cases it can be; please make sure that the override is completed.',
                    {'height':70,'width':500})
                worksheet.write_comment('J1',
                    'LSC or CSR No 4000: These are cases for which the funding code is “4000 LSC” but are marked as “LSC or CSR No.” With the exception of untimely closings, we should probably switch the funding codes here.',
                    {'height':70,'width':500})
                worksheet.write_comment('K1',
                    'No Age for Client: This report is of cases where no date of birth is reported. Please review the documents in the case file to see if there is an age or DOB on one of them that didn’t get entered into the correct fields. It is acceptable to put in an approximate date of birth and estimated age.',
                    {'height':70,'width':500})
                worksheet.write_comment('L1',
                    'Untimely closed: These are cases where the date of closing is long past the date opened. For example, a case opened in November of 2017 and closed as A (Counsel and Advice) or B (Brief Service) in April of 2019. All A/B cases should be closed in the year that they were opened, unless they were opened after October 1st of the year or after, unless there is a good reason. If a case was closed A/B outside that rule, there must be an override with a documented reason. Higher levels of service can be closed in the calendar year after the work is completed.',
                    {'height':70,'width':500})
                worksheet.write_comment('M1',
                    'Untimely Closed Overridden: Please review the reason for the override and make sure it is legitimate. Keep in mind that whether it is a “good reason” applies to us as a law firm and not the individual case handler. Reasons that are not legitimate include advocates not noticing the case had been assigned to them, or the file being lost/unassigned for a year and then an advocate closing it as soon it is assigned to them.',
                    {'height':70,'width':500})
                worksheet.write_comment('N1',
                    'Citizenship and Immigration Compliance: These are cases for which we should have the immigrant/citizenship compliance completed but do not, such as those where we met the client in person or we provided more than Advice or Brief Service. If a client is eligible under the anti-abuse statutes and all we need is a description of the basis of eligibility, the case note acts as verification. When you fix these, please remember to edit the closing information so that the CSR calculation is updated.',
                    {'height':70,'width':500})
                worksheet.write_comment('O1',
                    'Active Advocate Tester: These are cases for which the Primary Advocate does not have an active user profile in LegalServer. All Cases must be assigned to a currently-active case handler',
                    {'height':70,'width':500})
                    
                worksheet.conditional_format('D2:BO100000',{'type': 'text',
                                                         'criteria': 'containing',
                                                         'value': 'Needs',
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
    <title>Compliance Consolidater</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Summary of Compliance Reports</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Comply!!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2305" target="_blank">Compliance Consolidated Report</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
