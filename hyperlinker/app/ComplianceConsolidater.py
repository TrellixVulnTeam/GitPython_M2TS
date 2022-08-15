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
        
        #Checkbox determines which unit gets prepared, and offers alternative code for 'year end' compliance reports. 
        if request.form.get('LSU'):
            df = df[df['Assigned Branch/CC'] == "Legal Support Unit"]
        elif request.form.get('QLS'):
            df = df[df['Assigned Branch/CC'] == "Queens Legal Services"]
        elif request.form.get('MLS'):
            df = df[df['Assigned Branch/CC'] == "Manhattan Legal Services"]
        elif request.form.get('BLS'):
            df = df[df['Assigned Branch/CC'] == "Brooklyn Legal Services"]
        elif request.form.get('BxLS'):
            df = df[df['Assigned Branch/CC'] == "Bronx Legal Services"]
        elif request.form.get('SILS'):
            df = df[df['Assigned Branch/CC'] == "Staten Island Legal Services"]
        else:
            df = df
        #year end only looks at closed cases 
        if request.form.get('YearEnd'):
            df = df[df['Date Closed'] != ""]
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)     
        
        #Remove duplicate Case ID #s
        
        df = df.drop_duplicates(subset=['Hyperlinked CaseID#'], keep = 'first')
        
        #This is where all the functions happen:
        
        
        #No Legal Assistance Documented Compliance Check
        #Is "CSR CSR: Is Legal Assistance Document? = No"
        
        def NoAssistanceTester(NoAssistance):
            if NoAssistance == "No":
                return "Needs Review"
            else:
                return ''
                 
        df['No Legal Assistance Documented Tester'] = df.apply(lambda x : NoAssistanceTester(x['CSR: Is Legal Assistance Documented?']),axis=1)
        
        #No Time Entered for 90 Days
        #just open cases
        #is 'Age in Days' larger than 90? (need to de-duplicate to get just cases with most recent service date?)
        #don't do at year end
        
        if request.form.get('YearEnd'):
            print('Year End')
        else:
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
        
        
        def CitImmTester(CitizenshipStatus,AttestationOnFile,StaffVerified,ClientInPerson,CloseReason,ComplianceCheck):
            CloseReason = str(CloseReason)
            
            if AttestationOnFile == "Yes" and CitizenshipStatus == "Citizen":
                return ''
            elif CitizenshipStatus == "Non-Citizen" and StaffVerified == "Yes":
                return ''
            elif ComplianceCheck == "Yes":
                return ''
            elif CloseReason.startswith(('A','B')) == True and ClientInPerson == 'No':
                return ''
            else:
                return 'Needs Review'
            
            
        df['Citizenship & Immigration Tester'] = df.apply(lambda x : CitImmTester(x['Citizenship Status'],x['Attestation on File?'],x['Staff Verified Non-Citizenship Documentation'],x['Did any Staff Meet Client in Person?'],x['Close Reason'],x['Compliance Check Citizenship and Immigration']),axis=1)
        
        #Active Advocate Tester
        #not at year end
        if request.form.get('YearEnd'):
            print('Year End')
        else:
            def ActiveAdvocateTester(ActiveAdvocate):
                if ActiveAdvocate != "Yes":
                    return "Needs Review"
                else:
                    return ''
                 
            df['Active Advocate Tester'] = df.apply(lambda x : ActiveAdvocateTester(x['Login Active']),axis=1)
        
        #Case needs a retainer if it's closed with higher level of service than A or B, or has level of service entered, or is more than 30 days old (PAI cases don't need it)
        
        df['Today'] = datetime.now()
        df['TodayConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Today']), axis=1)
        df['OpenConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']), axis = 1)
        df['ConstructAge'] = df['TodayConstruct'] - df['OpenConstruct']
        
        def RetainerTester(Retainer,CloseReason,LevelOfService,PAICase,ConstructAge):
            if Retainer == 'Yes' or PAICase == 'Yes':
                return ''
            elif CloseReason.startswith('A') == True or CloseReason.startswith('B') == True:
                return ''
            elif CloseReason.startswith('F') == True or CloseReason.startswith('G') == True or CloseReason.startswith('H') == True or CloseReason.startswith('IA') == True or CloseReason.startswith('IB') == True or CloseReason.startswith('IC') == True or CloseReason.startswith('L') == True:
                return 'Needs Retainer'
            elif LevelOfService.startswith('Representation')== True:
                return 'Needs Retainer'
            elif LevelOfService == '' and ConstructAge > 200 and CloseReason == '':
                return 'Needs Retainer'
            else:
                return ''
        df['Retainer Tester'] = df.apply(lambda x : RetainerTester(x['Retainer on File'],x['Close Reason'],x['Level of Service'],x['PAI Case?'],x['ConstructAge']),axis=1)
        
        #Can't be closed before opened
        
        df['Opened Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']),axis=1)
        df['Closed Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Closed']),axis=1)
        
        def ClosedBeforeOpenedTester(DateOpened,DateClosed):
            if DateClosed == '':
                return ''
            elif DateClosed < DateOpened:
                return 'Needs Review'
            else:
                return ''
        df['Closed Before Opened Tester'] = df.apply(lambda x: ClosedBeforeOpenedTester(x['Opened Construct'],x['Closed Construct']),axis=1)        
                
        
        def PythonCSRTester(DateClosed,LSCTester,LegalAssistTester,UntimelyTester,CitImmTester):
            if DateClosed == '':
                return "N/A Not Closed"
            elif LSCTester != 'Yes' or 'Needs' in LegalAssistTester or 'Needs' in UntimelyTester or 'Needs' in CitImmTester:
                return 'No'
            else:
                return 'Yes'
        df['Python CSR Tester'] = df.apply(lambda x : PythonCSRTester(x['Date Closed'],x['LSC Eligible?'],x['No Legal Assistance Documented Tester'],x['Untimely Closed Tester'],x['Citizenship & Immigration Tester']),axis=1)
        
        def CSRAgreementTester (DateClosed,LegalServerCSR,PythonCSR):
            if DateClosed == '':
                return ''
            elif LegalServerCSR == 'Yes' and PythonCSR == 'Yes':
                return ''
            elif LegalServerCSR == 'No' and PythonCSR == 'No':
                return ''
            elif LegalServerCSR == 'Yes' and PythonCSR == 'No':
                return 'Needs Review'
            elif LegalServerCSR == 'No' and PythonCSR == 'Yes':
                return 'Needs Review'
            else:
                return 'How?!'
                
        df['CSR Agreement Tester'] = df.apply(lambda x : CSRAgreementTester(x['Date Closed'],x['CSR Eligible'],x['Python CSR Tester']),axis=1)
        
        
        # can't have zz IOLA outcome
        
        def OutcomeTester(Outcome):
            Outcome = str(Outcome)
            if 'ZZ-' in Outcome:
                return 'Needs Review'
            else: 
                return ''
        df['Outcome Tester'] = df.apply(lambda x : OutcomeTester(x['Outcome']),axis = 1)
        
        
        #make a tester tester (for year end)
        
        if request.form.get('YearEnd'):
            def TesterTester(NoAssistanceTester,TwoHundredPercentTester,OneTwentyFivePercentTester,LSCCSRTester,NoAgeTester,UntimelyClosedTester,UntimelyClosedOverriddenTester,CitImmTester,RetainerTester,ClosedBeforeOpenedTester,CSRAgreementTester,CaseStatus,OutcomeTester):
                if NoAssistanceTester != "" or TwoHundredPercentTester != "" or OneTwentyFivePercentTester != "" or LSCCSRTester != "" or NoAgeTester != "" or CitImmTester != "" or RetainerTester != "" or ClosedBeforeOpenedTester != "" or CSRAgreementTester != "" or OutcomeTester != "":
                    return "Needs Review"
                else:
                    return ""
            df['Tester Tester'] = df.apply(lambda x : TesterTester(x['No Legal Assistance Documented Tester'],x['200% of Poverty Tester'],x['125-200% of Poverty Tester'],x['Funding Code 4000 Tester'],x['No Age for Client Tester'],x['Untimely Closed Tester'],x['Untimely Closed Overridden Tester'],x['Citizenship & Immigration Tester'],x['Retainer Tester'],x['Closed Before Opened Tester'],x['CSR Agreement Tester'],x['Case Status'],x['Outcome Tester']),axis=1)        

        else:
        #includes active advocate and 90-day tester
            def TesterTester(NoAssistanceTester,AgeTester,TwoHundredPercentTester,OneTwentyFivePercentTester,LSCCSRTester,NoAgeTester,UntimelyClosedTester,UntimelyClosedOverriddenTester,CitImmTester,ActAdvTester,RetainerTester,ClosedBeforeOpenedTester,CSRAgreementTester,CaseStatus, OutcomeTester):
                if NoAssistanceTester != "" or AgeTester != "" or TwoHundredPercentTester != "" or OneTwentyFivePercentTester != "" or LSCCSRTester != "" or NoAgeTester != "" or UntimelyClosedTester != "" or UntimelyClosedOverriddenTester != "" or CitImmTester != "" or ActAdvTester != "" or RetainerTester != "" or ClosedBeforeOpenedTester != "" or CSRAgreementTester != "" or OutcomeTester != "":
                    return "Needs Review"
                else:
                    return ""
            df['Tester Tester'] = df.apply(lambda x : TesterTester(x['No Legal Assistance Documented Tester'],x['No Time Entered for 90 Days Tester'],x['200% of Poverty Tester'],x['125-200% of Poverty Tester'],x['Funding Code 4000 Tester'],x['No Age for Client Tester'],x['Untimely Closed Tester'],x['Untimely Closed Overridden Tester'],x['Citizenship & Immigration Tester'],x['Active Advocate Tester'],x['Retainer Tester'],x['Closed Before Opened Tester'],x['CSR Agreement Tester'],x['Case Status'],x['Outcome Tester']),axis=1)        
        
        #remove cases that don't need compliance review
        df = df[df['Tester Tester'] != ""]
        df = df[df['Assigned Branch/CC'] != "Legal Services NYC"]
        
        #change name of column for intake worker
        df['Intake Caseworker'] = df['Caseworker Name']
        
        #Putting everything in the right order
        
        
        if request.form.get('YearEnd'):
            df = df[['Hyperlinked CaseID#',
            'Assigned Branch/CC',
            'Primary Advocate Name',
            'Client First Name',
            'Client Last Name',
            'No Legal Assistance Documented Tester',
            '200% of Poverty Tester',
            '125-200% of Poverty Tester',
            'Funding Code 4000 Tester',
            'No Age for Client Tester',
            'Citizenship & Immigration Tester',
            'Retainer Tester',
            'Closed Before Opened Tester',
            'CSR Agreement Tester',
            'Outcome Tester',
            'Date Opened',
            'Case Status',
            'Level of Service',
            'Intake Caseworker',
            'Legal Problem Code',
            'Compliance Check Untimely Closed',
            'Compliance Check Untimely Closed Overwritten',
            'Compliance Check 125 to 200 Poverty Income Ineligible',
            'Compliance Check 200 Poverty Income Eligible',
            'Compliance Check Citizenship and Immigration',
            'Citizenship Status',
            'Attestation on File?',
            'Staff Verified Non-Citizenship Documentation',
            'Did any Staff Meet Client in Person?',
            'Close Reason',
            'Date Closed',
            'CSR: Is Legal Assistance Documented?',
            'Age in Days',
            'Percentage of Poverty',
            'LSC Eligible?',
            'CSR Eligible',
            'Income Eligible',
            'Primary Funding Codes',
            'Secondary Funding Codes',
            'DOB Information',
            'Group',
            'CSR: Timely Closing?',
            'Was Timely Closed overridden?',
            'PAI Case?',
            'Retainer on File',
            'Python CSR Tester',
            'TodayConstruct',
            'ConstructAge',
            'Special Legal Problem Code',
            'Tester Tester',
            'Outcome',
            
            'Most Recent Note',
            'Serv Max Body']]
            
        #not year end version
        else:
            df = df[['Hyperlinked CaseID#',
            'Assigned Branch/CC',
            'Primary Advocate Name',
            'Client First Name',
            'Client Last Name',
            'No Legal Assistance Documented Tester',
            'No Time Entered for 90 Days Tester',
            '200% of Poverty Tester',
            '125-200% of Poverty Tester',
            'Funding Code 4000 Tester',
            'No Age for Client Tester',
            'Untimely Closed Tester',
            'Untimely Closed Overridden Tester',
            'Citizenship & Immigration Tester',
            'Active Advocate Tester',
            'Retainer Tester',
            'Closed Before Opened Tester',
            'CSR Agreement Tester',
            'Outcome Tester',
            'Date Opened',
            'Case Status',
            'Level of Service',
            'Intake Caseworker',
            'Legal Problem Code',
            'Compliance Check Untimely Closed',
            'Compliance Check Untimely Closed Overwritten',
            'Compliance Check 125 to 200 Poverty Income Ineligible',
            'Compliance Check 200 Poverty Income Eligible',
            'Compliance Check Citizenship and Immigration',
            'Citizenship Status',
            'Attestation on File?',
            'Staff Verified Non-Citizenship Documentation',
            'Did any Staff Meet Client in Person?',
            'Close Reason',
            'Date Closed',
            'CSR: Is Legal Assistance Documented?',
            'Age in Days',
            'Percentage of Poverty',
            'LSC Eligible?',
            'CSR Eligible',
            'Income Eligible',
            'Primary Funding Codes',
            'Secondary Funding Codes',
            'DOB Information',
            'Group',
            'CSR: Timely Closing?',
            'Was Timely Closed overridden?',
            'PAI Case?',
            'Retainer on File',
            'Python CSR Tester',
            'TodayConstruct',
            'ConstructAge',
            'Special Legal Problem Code',
            'Tester Tester',
            'Outcome',
            
            'Most Recent Note',
            'Serv Max Body']]
        
        #LSU Splitter:
        
        ProBonoList = [
        "Cardenas, Lizeth",
        "Chua, Janice W",
        "Heintz, Adam J",
        "McCormick, James H",
        "Morales-Robinson, Ana Y.",
        "Sahai, Chelsea E",
        "Collins, Lisa",
        "Ruiz-Caceres, Gaby A",
        "Harrington, Elizabeth K",
        "Guzina, Ana"]
        
        VeronicaList = ["Cook, Veronica J"]
        EdList = ["Josephson, Edward"]
        ImmigrationList = ["Chen, Rex J"]
        BankruptcyList = ["Kransdorf, William Z","Miranda, Stephanie"]
        IntakeList = ["Eagan, Emilie A."]
        
        def TabSplitter(PrimaryAdvocate):
            if PrimaryAdvocate in ProBonoList:
                return "Pro Bono"
            elif PrimaryAdvocate in EdList:
                return "Ed"
            elif PrimaryAdvocate in ImmigrationList:
                return "Immigration"
            elif PrimaryAdvocate in BankruptcyList:
                return "Bankruptcy"
            elif PrimaryAdvocate in IntakeList:
                return "Intake"
            elif PrimaryAdvocate in VeronicaList:
                return "Veronica"
            else:
                return "Not LSU"
        
        #Choose what to split things by:
        if request.form.get('LSU'):
            df['TabSplitValue'] = df.apply(lambda x : TabSplitter(x['Primary Advocate Name']),axis = 1)
        elif request.form.get('BLSbyAdvocate'):
            df['TabSplitValue'] = df['Primary Advocate Name']
        else:
            df['TabSplitValue'] = df.apply(lambda x : DataWizardTools.UnitSplitter(x['Legal Problem Code']),axis = 1)
        
        df = df.sort_values(by=['Date Opened'])
        df = df.sort_values(by=['Primary Advocate Name'])
        
        #Preparing Excel Document

        if request.form.get('MLSByTester'):
            NoLegalAssistance_df = df[df['No Legal Assistance Documented Tester'] == "Needs Review"]
            NoTimeEntered_df = df[df['No Time Entered for 90 Days Tester'] == "Needs Review"]
            TwoHundredPoverty_df = df[df['200% of Poverty Tester'] == "Needs Review"]
            OneTwentyFivePovertydf = df[df['125-200% of Poverty Tester'] == "Needs Review"]
            FundingCode4000_df = df[df['Funding Code 4000 Tester'] == "Needs Review"]
            NoAge_df = df[df['No Age for Client Tester'] == "Needs Review"]
            UntimelyClosed_df = df[df['Untimely Closed Tester'] == "Needs Review"]
            UntimelyOverridden_df = df[df['Untimely Closed Overridden Tester'] == "Needs Review"]
            Citizenship_df = df[df['Citizenship & Immigration Tester'] == "Needs Review"]
            ActiveAdvocate_df = df[df['Active Advocate Tester'] == "Needs Review"]
            Retainer_df = df[df['Retainer Tester'] == "Needs Review"]
            
        
            borough_dictionary = {'No Legal Assistance' : NoLegalAssistance_df, 'No Time Entered' : NoTimeEntered_df, '200% Poverty': TwoHundredPoverty_df , '125% Poverty': OneTwentyFivePovertydf , 'Funding Code 4000': FundingCode4000_df , 'No Age': NoAge_df , 'Untimely Closed': UntimelyClosed_df , 'Untimely Overridden': UntimelyOverridden_df , 'Citizenship': Citizenship_df , 'Active Advocate': ActiveAdvocate_df , 'Retainer': Retainer_df} 
            
            print(borough_dictionary)
        else:
            borough_dictionary = dict(tuple(df.groupby('TabSplitValue')))
            
            
            
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
                'valign': 'top',
                'align': 'center',
                'bg_color' : '#eeece1'
                })
                
                
                worksheet = writer.sheets[i]
                
                #Add column header data back in
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                worksheet.autofilter('A1:V1')
                worksheet.set_column('A:A',15,link_format)
                worksheet.set_column('C:C',20)
                worksheet.set_column('D:W',13)
                worksheet.set_column('B:B',0)
                worksheet.set_column('X:ZZ',0)
                if request.form.get('YearEnd'):
                    worksheet.set_column('S:ZZ',0)
                worksheet.set_row(0,60)
                worksheet.write_comment('F1',
                    'No Legal Assistance Documented: If we close a case as something other than ZZ, we must have some documentation of the legal assistance provided. For these, whoever closed the case may have misunderstood the meaning of “legal assistance” or “documented.” Please review and add the legal assistance documented in case notes. If there was no legal assistance, change the closing code to ZZ.',
                    {'height':70,'width':500,'x_offset':-30})
                
                NoTimeInNinetyDays = 'No Time in 90 Days: These are cases for which no time has been entered in 90 days. Please make sure these are all active. If they are active and pending in a court or other forum, enter a case note stating this. We have come across cases opened as long as four or five years ago for which no time had been entered since the year of opening, and which should have been closed in the same year they were opened.'
                
                if request.form.get('YearEnd'):
                    print('Year End')
                else:
                    worksheet.write_comment('G1',
                       NoTimeInNinetyDays,
                        {'height':70,'width':500})
                
                TwoHundredPercent = '200% of Poverty Income Eligible: Please confirm that the overrides were completed properly in these cases; these are rarely correct, but it is possible. If they were properly done, there is no need to do anything else.'
                
                if request.form.get('YearEnd'):
                    worksheet.write_comment('G1',
                    TwoHundredPercent,
                    {'height':70,'width':500})
                else:
                    worksheet.write_comment('H1',
                        TwoHundredPercent,
                        {'height':70,'width':500})
                        
                    
                OneTwentyFiveToTwoHundred = '125-200 Poverty Income Ineligible with Type of Income: These are cases where the client is between 125% and 200% of poverty and the financial override was not completed. In most cases it can be; please make sure that the override is completed.'    
                

                if request.form.get('YearEnd'):
                    worksheet.write_comment('H1',
                        OneTwentyFiveToTwoHundred,
                        {'height':70,'width':500})
                else:                
                    worksheet.write_comment('I1',
                        OneTwentyFiveToTwoHundred,
                        {'height':70,'width':500})
                    
                    
                LSCorCSRnoFourThousand = 'LSC or CSR No 4000: These are cases for which the funding code is “4000 LSC” but are marked as “LSC or CSR No.” With the exception of untimely closings, we should probably switch the funding codes here.'

                if request.form.get('YearEnd'):
                    worksheet.write_comment('I1',
                        LSCorCSRnoFourThousand,
                        {'height':70,'width':500})
                else:
                    worksheet.write_comment('J1',
                        LSCorCSRnoFourThousand,
                        {'height':70,'width':500})
                    
                
                NoAgeForClient= 'No Age for Client: This report is of cases where no date of birth is reported. Please review the documents in the case file to see if there is an age or DOB on one of them that didn’t get entered into the correct fields. It is acceptable to put in an approximate date of birth and estimated age.'
                 
                if request.form.get('YearEnd'): 
                    worksheet.write_comment('J1',
                        NoAgeForClient,
                        {'height':70,'width':500})
                else:
                    worksheet.write_comment('K1',
                        NoAgeForClient,
                        {'height':70,'width':500})
                
                
                UnTimelyClosed = 'Untimely closed: These are cases where the date of closing is long past the date opened. For example, a case opened in November of 2017 and closed as A (Counsel and Advice) or B (Brief Service) in April of 2019. All A/B cases should be closed in the year that they were opened, unless they were opened after October 1st of the year or after, unless there is a good reason. If a case was closed A/B outside that rule, there must be an override with a documented reason. Higher levels of service can be closed in the calendar year after the work is completed.'
                
                if request.form.get('YearEnd'): 
                    print('MLS Custom')
                else:
                    worksheet.write_comment('L1',
                        UnTimelyClosed,
                        {'height':70,'width':500})
                
                
                
                UnTimelyOverRidden = 'Untimely Closed Overridden: Please review the reason for the override and make sure it is legitimate. Keep in mind that whether it is a “good reason” applies to us as a law firm and not the individual case handler. Reasons that are not legitimate include advocates not noticing the case had been assigned to them, or the file being lost/unassigned for a year and then an advocate closing it as soon it is assigned to them.'
                
                if request.form.get('YearEnd'): 
                    print('MLS Custom')
                else:
                    worksheet.write_comment('M1',
                        UnTimelyOverRidden,
                        {'height':70,'width':500})
                
                
                CitizenAndImmigration = 'Citizenship and Immigration Compliance: These are cases for which we should have the immigrant/citizenship compliance completed but do not, such as those where we met the client in person or we provided more than Advice or Brief Service. If a client is eligible under the anti-abuse statutes and all we need is a description of the basis of eligibility, the case note acts as verification. When you fix these, please remember to edit the closing information so that the CSR calculation is updated.'
                if request.form.get('YearEnd'):     
                    worksheet.write_comment('K1',
                        CitizenAndImmigration,
                        {'height':70,'width':500})
                else: 
                    worksheet.write_comment('N1',
                        CitizenAndImmigration,
                        {'height':70,'width':500})
                
                ActiveAdvocateText = 'Active Advocate Tester: These are cases for which the Primary Advocate does not have an active user profile in LegalServer. All Cases must be assigned to a currently-active case handler'
                
                if request.form.get('YearEnd'):   
                    print('Year End')
                else: 
                    worksheet.write_comment('O1',
                        ActiveAdvocateText,
                        {'height':70,'width':500})
                
                
                RetainerText = 'Retainer Tester: Cases that have a Close Reason showing more than Advice/Brief Service, or a Level of Service indicating representation has been provided must have retainers. Open cases with a blank Level of Service will ask for retainers after they have been open for 2 months.'
                
                if request.form.get('YearEnd'):  
                    worksheet.write_comment('L1',
                        RetainerText,
                        {'height':70,'width':500})
                else: 
                    worksheet.write_comment('P1',
                        RetainerText,
                        {'height':70,'width':500})
                
                ClosedBeforeOpenedText = 'Closed Before Opened Tester: The date that a case was closed must be equal to or later than the date on which it was opened.'
                
                if request.form.get('YearEnd'): 
                    worksheet.write_comment('M1',
                        ClosedBeforeOpenedText,
                        {'height':70,'width':500})
                else: 
                    worksheet.write_comment('Q1',
                        ClosedBeforeOpenedText,
                        {'height':70,'width':500})
                
                CSRAgreementText = 'CSR Agreement Tester: The underlying case data seems to contradict the CSR value calculated by LegalServer. Please review this case to ensure that CSR was recalculated appropriately or that a manual override was performed for valid reasons.'
                
                if request.form.get('YearEnd'): 
                    worksheet.write_comment('N1',
                        CSRAgreementText,
                        {'height':70,'width':500})
                else:
                    worksheet.write_comment('R1',
                        CSRAgreementText,
                        {'height':70,'width':500})

                OutcomeText = 'This case has been assigned an IOLA outcome of "ZZ-Administrative Closing", but has a close reason indicating that services were provided. Please review and correct this inconsistency.'
                
                if request.form.get('YearEnd'): 
                    worksheet.write_comment('O1',
                        OutcomeText,
                        {'height':70,'width':500})
                else:
                    worksheet.write_comment('S1',
                        OutcomeText,
                        {'height':70,'width':500})



                
                worksheet.conditional_format('D2:V100000',{'type': 'text',
                                                         'criteria': 'containing',
                                                         'value': 'Needs',
                                                         'format': problem_format})
                worksheet.freeze_panes(1,1)
            writer.save()
        output_filename = f.filename

        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)

        if request.form.get('MLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "MLS " + f.filename)
        elif request.form.get('MLSByTester'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "MLS by Tester" + f.filename)
        elif request.form.get('BLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "BLS " + f.filename)
        elif request.form.get('BLSbyAdvocate'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "BLS by Advocate" + f.filename)
        elif request.form.get('BxLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "BxLS " + f.filename)
        elif request.form.get('SILS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "SILS " + f.filename)
        elif request.form.get('QLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "QLS " + f.filename)
        elif request.form.get('LSU'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "LSU " + f.filename)
        else:
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "All Boroughs " + f.filename)

        
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
    
    </br>
    </br>
    <input type="checkbox" id="LSU" name="LSU" value="LSU">
    <label for="LSU"> LSU Compliance</label><br>
    <input type="checkbox" id="QLS" name="QLS" value="QLS">
    <label for="QLS"> QLS Compliance</label><br>
    <input type="checkbox" id="MLS" name="MLS" value="MLS">
    <label for="MLS"> MLS Compliance</label><br>
    <input type="checkbox" id="BLS" name="BLS" value="BLS">
    <label for="BLS"> BLS Compliance</label><br>
    <input type="checkbox" id="MLSByTester" name="MLSByTester" value="MLSByTester">
    <label for="MLSByTester"> MLS Compliance by Tester</label><br>
    <input type="checkbox" id="BLSbyAdvocate" name="BLSbyAdvocate" value="BLSbyAdvocate">
    <label for="BLSbyAdvocate"> BLS Compliance by Advocate</label><br>
    <input type="checkbox" id="BxLS" name="BxLS" value="BxLS">
    <label for="BxLS"> BxLS Compliance</label><br>
    <input type="checkbox" id="SILS" name="SILS" value="SILS">
    <label for="SILS"> SILS Compliance</label><br>
    <input type="checkbox" id="YearEnd" name="YearEnd" value="YearEnd">
    <label for="YearEnd"> Year End?</label><br>
    
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2305" target="_blank">Compliance Consolidated Report</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
