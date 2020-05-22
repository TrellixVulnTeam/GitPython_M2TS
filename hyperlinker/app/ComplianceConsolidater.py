#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
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
        
        #Import Excel Sheet
        data_xls = pd.read_excel(f,skiprows=2)
        
        #delete any rows that don't have case ID#s in them
        
        def NoIDDelete(CaseID):
            if CaseID == '' or CaseID == 'nan':
                return 'No Case ID'
            else:
                return str(CaseID)
        data_xls['Matter/Case ID#'] = data_xls.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        
        data_xls = data_xls[data_xls['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[-7:])
        data_xls['Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + data_xls['Matter/Case ID#'] +'"' +')'
        move = data_xls['Hyperlinked Case #']
        del data_xls['Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #',move)           
        
        #Remove duplicate Case ID #s
        
        #data_xls = data_xls.drop_duplicates(subset='Hyperlinked Case #', keep = 'first')
        
        #This is where all the functions happen:
        
        
        #No Legal Assistance Documented Compliance Check
        #Is Case Closed?
        #Is "CSR CSR: Is Legal Assistance Document? = No"
        
        def NoAssistanceTester(DateClosed,NoAssistance):
            if DateClosed != "nan" and NoAssistance == "No":
                return "Needs Review"
            else:
                return ''
                 
        data_xls['No Legal Assistance Documented Tester'] = data_xls.apply(lambda x : NoAssistanceTester(x['Date Closed'],x['CSR: Is Legal Assistance Documented?']),axis=1)
        
        #No Time Entered for 90 Days
        #just open cases
        #is 'Age in Days' larger than 90? (need to de-duplicate to get just cases with most recent service date?)
        
        def AgeTester(DateClosed,AgeInDays):
            DateClosed = str(DateClosed)
            if DateClosed != "nan" and AgeInDays > 90:
                return 'Needs Review'
            else:
                return ''
        
        data_xls['No Time Entered for 90 Days Tester'] = data_xls.apply(lambda x : AgeTester(x['Date Closed'],x['Age In Days']),axis=1)
        
        
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
        
        data_xls['200% of Poverty Tester'] = data_xls.apply(lambda x : TwoHundredPercentTester(x['Percentage of Poverty'],x['LSC Eligible?'],x['CSR Eligible'],x['Compliance Check 200 Poverty Income Eligible']),axis=1)
        
        
        #125-200 poverty income eligible
        #percentage of poverty is between 125 and 200 AND 
        #Income eligible = no
        
        def OneTwentyFivePercentTester(PovPercent,IncomeEligible, ComplianceCheck):
            if PovPercent > 125 and PovPercent < 200 and IncomeEligible == 'No' and ComplianceCheck != "Yes":
                return "Needs Review"
            else:
                return ''
        
        data_xls['125-200% of Poverty Tester'] = data_xls.apply(lambda x : OneTwentyFivePercentTester(x['Percentage of Poverty'],x['Income Eligible'],x['Compliance Check 125 to 200 Poverty Income Ineligible']),axis=1)
        
        
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

        data_xls['Funding Code 4000 Tester'] = data_xls.apply(lambda x : LSCCSRTester(x['Primary Funding Codes'],x['Secondary Funding Codes'],x['LSC Eligible?'],x['CSR Eligible']),axis=1)
        
        #No Age for Client
        #Not a Group
        #DOB Status = Refused or Unknown (approximate or blank are fine)
        
        def NoAgeTester (DOBStatus,Group):
            if DOBStatus == 101129 and Group != "Yes":
                return "Needs Review"
            else:
                return ""
        
        data_xls['No Age for Client Tester'] = data_xls.apply(lambda x : NoAgeTester(x['DOB Status [lookup]'],x['Group']),axis=1)
        
        #Untimely Closed
        #Closed Cases
        #CSR: Timely Closing = no
        
        def UntimelyClosedTester(DateClosed,TimelyClosing,ComplianceCheck):
            if DateClosed != "nan" and TimelyClosing == "No" and ComplianceCheck != "Yes":
                return "Needs Review"
            else:
                return ''
                 
        data_xls['Untimely Closed Tester'] = data_xls.apply(lambda x : UntimelyClosedTester(x['Date Closed'],x['CSR: Timely Closing?'],x['Compliance Check Untimely Closed']),axis=1)
        
        #Untimely Closed Overridden
        #is case closed?
        #CSR Timely Closing = YES AND
        #Was timely closed overridden = yes
        
        
        def UntimelyClosedOverriddenTester(DateClosed,TimelyClosing,TimelyClosingOverridden):
            if DateClosed != "nan" and TimelyClosing == "Yes" and TimelyClosingOverridden == "Yes":
                return "Needs Review"
            else:
                return ''
                 
        data_xls['Untimely Closed Overridden Tester'] = data_xls.apply(lambda x : UntimelyClosedOverriddenTester(x['Date Closed'],x['CSR: Timely Closing?'],x['Was Timely Closed overridden?']),axis=1)

        
        #Citizenship and Immigration Compliance
        #Is case closed?
        #is attestation on file = no or unanswered
        #and staff verified non-citizenship documentation = no or unanswered
        #AND either
        #did staff meet client in person = yes or unanswered
        #or close reason starts with = F,G,H,I,or L
        
        
        def CitImmTester(AttestationOnFile,StaffVerified,ClientInPerson,CloseReason,ComplianceCheck):
            CloseReason = str(CloseReason)
            if CloseReason != "nan"and AttestationOnFile != "Yes" and StaffVerified != "Yes" and ClientInPerson != 'No' and ComplianceCheck != "Yes":
                return 'Needs Review'
            
            elif AttestationOnFile != "Yes" and StaffVerified != "Yes" and CloseReason.startswith(('F','G','H','I','L')) == True and ComplianceCheck != "Yes": 
                return 'Needs Review'
                
            else:
                return ''                

        data_xls['Citizenship & Immigration Tester'] = data_xls.apply(lambda x : CitImmTester(x['Attestation on File?'],x['Staff Verified Non-Citizenship Documentation'],x['Did any Staff Meet Client in Person?'],x['Close Reason'],x['Compliance Check Citizenship and Immigration']),axis=1)
        
        
        #also add in code related to checking the boxes they check if something has been compliance-reviewed
        
        #Make a tester tester
        def TesterTester(NoAssistanceTester,AgeTester,TwoHundredPercentTester,OneTwentyFivePercentTester,LSCCSRTester,NoAgeTester,UntimelyClosedTester,UntimelyClosedOverriddenTester,CitImmTester):
            if NoAssistanceTester != "" or AgeTester != "" or TwoHundredPercentTester != "" or OneTwentyFivePercentTester != "" or LSCCSRTester != "" or NoAgeTester != "" or UntimelyClosedTester != "" or UntimelyClosedOverriddenTester != "" or CitImmTester != "":
                return "Needs Review"
            else:
                return ""
        data_xls['Tester Tester'] = data_xls.apply(lambda x : TesterTester(x['No Legal Assistance Documented Tester'],x['No Time Entered for 90 Days Tester'],x['200% of Poverty Tester'],x['125-200% of Poverty Tester'],x['Funding Code 4000 Tester'],x['No Age for Client Tester'],x['Untimely Closed Tester'],x['Untimely Closed Overridden Tester'],x['Citizenship & Immigration Tester']),axis=1)        
        
        #remove cases that don't need compliance review
        data_xls = data_xls[data_xls['Tester Tester'] == "Needs Review"]
        
        
        ##split by borough into different tabs?
        
        data_xls = data_xls[['Hyperlinked Case #','Assigned Branch/CC','Primary Advocate','Client Name','No Legal Assistance Documented Tester','No Time Entered for 90 Days Tester','200% of Poverty Tester','125-200% of Poverty Tester','Funding Code 4000 Tester','No Age for Client Tester','Untimely Closed Tester','Untimely Closed Overridden Tester','Citizenship & Immigration Tester','Compliance Check Untimely Closed','Compliance Check 125 to 200 Poverty Income Ineligible','Compliance Check 200 Poverty Income Eligible','Compliance Check Citizenship and Immigration','Attestation on File?','Staff Verified Non-Citizenship Documentation','Did any Staff Meet Client in Person?','Close Reason','Date Closed','CSR: Is Legal Assistance Documented?','Age In Days','Percentage of Poverty','LSC Eligible?','CSR Eligible','Income Eligible','Primary Funding Codes','Secondary Funding Codes','DOB Status [lookup]','Group','CSR: Timely Closing?','Was Timely Closed overridden?','Tester Tester']]
        
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Compliance Summary',index=False, header = False,startrow=1)
        workbook = writer.book
        worksheet = writer.sheets['Compliance Summary']
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        problem_format = workbook.add_format({'bg_color':'yellow'})
        header_format = workbook.add_format({
            'text_wrap':True,
            'bold':True,
            'valign': 'middle',
            'align': 'center'
            })
        
        #Add column header data back in
        for col_num, value in enumerate(data_xls.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        
        #assign new format to column A
        worksheet.set_column('A:A',15,link_format)
        worksheet.set_column('B:M',20)
        worksheet.set_column('N:ZZ',0)
        worksheet.set_row(0,60)
        worksheet.write_comment('E1',
            'No Legal Assistance Documented: If we close a case as something other than ZZ, we must have some documentation of the legal assistance provided. For these, whoever closed the case may have misunderstood the meaning of “legal assistance” or “documented.” Please review and add the legal assistance documented in case notes. If there was no legal assistance, change the closing code to ZZ.',
            {'height':70,'width':500,'x_offset':-30})
        worksheet.write_comment('F1',
            'No Time in 90 Days: These are cases for which no time has been entered in 90 days. Please make sure these are all active. If they are active and pending in a court or other forum, enter a case note stating this. We have come across cases opened as long as four or five years ago for which no time had been entered since the year of opening, and which should have been closed in the same year they were opened.',
            {'height':70,'width':500})
        worksheet.write_comment('G1',
            '200% of Poverty Income Eligible: Please confirm that the overrides were completed properly in these cases; these are rarely correct, but it is possible. If they were properly done, there is no need to do anything else.',
            {'height':70,'width':500})
        worksheet.write_comment('H1',
            '125-200 Poverty Income Ineligible with Type of Income: These are cases where the client is between 125% and 200% of poverty and the financial override was not completed. In most cases it can be; please make sure that the override is completed.',
            {'height':70,'width':500})
        worksheet.write_comment('I1',
            'LSC or CSR No 4000: These are cases for which the funding code is “4000 LSC” but are marked as “LSC or CSR No.” With the exception of untimely closings, we should probably switch the funding codes here.',
            {'height':70,'width':500})
        worksheet.write_comment('J1',
            'No Age for Client: This report is of cases where no date of birth is reported. Please review the documents in the case file to see if there is an age or DOB on one of them that didn’t get entered into the correct fields. It is acceptable to put in an approximate date of birth and estimated age.',
            {'height':70,'width':500})
        worksheet.write_comment('K1',
            'Untimely closed: These are cases where the date of closing is long past the date opened. For example, a case opened in November of 2017 and closed as A (Counsel and Advice) or B (Brief Service) in April of 2019. All A/B cases should be closed in the year that they were opened, unless they were opened after October 1st of the year or after, unless there is a good reason. If a case was closed A/B outside that rule, there must be an override with a documented reason. Higher levels of service can be closed in the calendar year after the work is completed.',
            {'height':70,'width':500})
        worksheet.write_comment('L1',
            'Untimely Closed Overridden: Please review the reason for the override and make sure it is legitimate. Keep in mind that whether it is a “good reason” applies to us as a law firm and not the individual case handler. Reasons that are not legitimate include advocates not noticing the case had been assigned to them, or the file being lost/unassigned for a year and then an advocate closing it as soon it is assigned to them.',
            {'height':70,'width':500})
        worksheet.write_comment('M1',
            'Citizenship and Immigration Compliance: These are cases for which we should have the immigrant/citizenship compliance completed but do not, such as those where we met the client in person or we provided more than Advice or Brief Service. If a client is eligible under the anti-abuse statutes and all we need is a description of the basis of eligibility, the case note acts as verification. When you fix these, please remember to edit the closing information so that the CSR calculation is updated.',
            {'height':70,'width':500})
            
        worksheet.conditional_format('D2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
        worksheet.freeze_panes(1,1)
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Python " + f.filename)
        
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
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2300" target="_blank">Compliance Data Cleanup Consolidated Report</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
