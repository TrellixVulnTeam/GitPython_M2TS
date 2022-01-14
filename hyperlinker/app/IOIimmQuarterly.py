from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, ImmigrationToolBox, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/IOIimmQuarterly", methods=['GET', 'POST'])
def upload_IOIimmQuarterly():
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
        
        df.fillna('',inplace=True)
       #Create Hyperlinks
        df['Hyperlinked Case #'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)
        
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Bronx Legal Services','BxLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Brooklyn Legal Services','BkLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Queens Legal Services','QLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Manhattan Legal Services','MLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Staten Island Legal Services','SILS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Legal Support Unit','LSU')
        
        
        
        #Determining 'level of service' from 3 fields       
          
        df['HRA Level of Service'] = df.apply(lambda x: ImmigrationToolBox.HRA_Level_Service(x['Close Reason'],x['Level of Service']), axis=1)
        
        
       #HRA Case Coding
       #Putting Cases into HRA's Baskets!  
        df['HRA Case Coding'] = df.apply(lambda x: ImmigrationToolBox.HRA_Case_Coding(x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA Level of Service'],x['IOI Does Client Have A Criminal History? (IOI 2)']), axis=1)

        #Dummy SLPC for Juvenile Cases
        def DummySLPC(LPC,SLPC):
            LPC = str(LPC)
            SLPC = str(SLPC)
            if LPC == "44 Minor Guardianship / Conservatorship" or LPC == "42 Neglected/Abused/Dependent":
                return 'N/A'
            else:
                return SLPC
                
        df['Special Legal Problem Code'] = df.apply(lambda x: DummySLPC(x['Legal Problem Code'],x['Special Legal Problem Code']), axis=1)
    
        df['HRA Service Type'] = df['HRA Case Coding'].apply(lambda x: x[:2] if x != 'Hold For Review' else '')

        df['HRA Proceeding Type'] = df['HRA Case Coding'].apply(lambda x: x[3:] if x != 'Hold For Review' else '')
        
        #Giving things better names in cleanup sheet
        
        df['DHCI form?'] = df['Has Declaration of Household Composition and Income (DHCI) Form?']
        
        df['Consent form?'] = df['IOI HRA Consent Form? (IOI 2)']
        
        df['Client Name'] = df['Full Person/Group Name (Last First)']
        
        df['Office'] = df['Assigned Branch/CC']
        
        df['Country of Origin'] = df['IOI Country Of Origin (IOI 1 and 2)']
        
        df['Substantial Activity'] = df['IOI Substantial Activity (Choose One)']
        
        df['Date of Substantial Activity'] = df['Custom IOI Date substantial Activity Performed']
        
        #Income Waiver
        
        def Income_Exclude(IncomePct,Waiver,Referral):   
            IncomePct = int(IncomePct)
            Waiver = str(Waiver)
            if Referral == 'Action NY':
                return ''
            elif IncomePct > 200 and Waiver.startswith('2') == False:
                return 'Needs Income Waiver'
            else:
                return ''

        df['Exclude due to Income?'] = df.apply(lambda x: Income_Exclude(x['Percentage of Poverty'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)'],x['IOI Referral Source (IOI 2)']), axis=1)
        
        #Eligibility_Date & Rollovers 
        
        def Eligibility_Date(Substantial_Activity_Date,Effective_Date,Date_Opened):
            if Substantial_Activity_Date != '':
                return Substantial_Activity_Date
            elif Effective_Date != '':
                return Effective_Date
            else:
                return Date_Opened
        df['Eligibility_Date'] = df.apply(lambda x : Eligibility_Date(x['IOI Date Substantial Activity Performed 2022'],x['IOI HRA Effective Date (optional) (IOI 2)'],x['Date Opened']), axis = 1)
        
    
        #Manipulable Dates               
        
        
        df['Open Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Eligibility_Date']),axis = 1)
        
        #Substantial Activity construct
        df['Subs Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['IOI Date Substantial Activity Performed 2022']),axis = 1)
        
        
        df['Outcome1 Month'] = df['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[:2])
        df['Outcome1 Day'] = df['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[3:5])
        df['Outcome1 Year'] = df['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[6:])
        df['Outcome1 Construct'] = df['Outcome1 Year'] + df['Outcome1 Month'] + df['Outcome1 Day']

        df['Outcome2 Month'] = df['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[:2])
        df['Outcome2 Day'] = df['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[3:5])
        df['Outcome2 Year'] = df['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[6:])
        df['Outcome2 Construct'] = df['Outcome2 Year'] + df['Outcome2 Month'] + df['Outcome2 Day']       
        
        
        #DHCI Form
                
        def DHCI_Needed(DHCI,Open_Construct,LoS):
            if Open_Construct == '':
                return ''
            elif LoS.startswith('Advice'):
                return ''
            elif LoS.startswith('Brief'):
                return ''
            elif int(Open_Construct) < 20180701:
                return ''
            elif DHCI != 'Yes':
                return 'Needs DHCI Form'
            else:
                return ''
        
        df['Needs DHCI?'] = df.apply(lambda x: DHCI_Needed(x['Has Declaration of Household Composition and Income (DHCI) Form?'],x['Open Construct'],x['Level of Service']), axis=1)
        
        
        
        #Needs Substantial Activity to Rollover into FY'22
        
        
        
        df['Needs Substantial Activity?'] = df.apply(lambda x: ImmigrationToolBox.Needs_Rollover(x['Open Construct'],x['IOI FY22 Substantial Activity 2022'],x['Subs Construct'],x['Matter/Case ID#']), axis=1) 
        
        


        #Outcomes
        
                
        #if there are two outcomes choose which outcome to report based on which one happened more recently 
                
        def OutcomeToReport (Outcome1,OutcomeDate1,Outcome2,OutcomeDate2,ServiceLevel,CloseDate):
            if OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Advice':
                return 'Advice given'
            elif OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Brief Service':
                return 'Advice given'
            elif CloseDate != '' and ServiceLevel == 'Full Rep or Extensive Service' and Outcome1 == '' and Outcome2 == '':
                return '*Needs Outcome*'
            elif OutcomeDate1 >= OutcomeDate2:
                return Outcome1
            elif OutcomeDate2 > OutcomeDate1:
                return Outcome2
            else:
                return 'no actual outcome'
        
        df['Outcome To Report'] = df.apply(lambda x: OutcomeToReport(x['IOI Outcome 2 (IOI 2)'],x['Outcome1 Construct'],x['IOI Secondary Outcome 2 (IOI 2)'],x['Outcome2 Construct'],x['HRA Level of Service'],x['Date Closed']), axis=1)
      
        #make it add the outcome date as well (or tell you if you need it!)        

        def OutcomeDateToReport (Outcome1,OutcomeDate1,Outcome2,OutcomeDate2,ServiceLevel,CloseDate,ActualOutcomeDate1,ActualOutcomeDate2):
            if OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Advice':
                return CloseDate
            elif OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Brief Service':
                return CloseDate           
            elif OutcomeDate1 >= OutcomeDate2:
                return ActualOutcomeDate1
            elif OutcomeDate2 > OutcomeDate1:
                return ActualOutcomeDate2
            else:
                return '*Needs Outcome Date*'
            
                
        df['Outcome Date To Report'] = df.apply(lambda x: OutcomeDateToReport(x['IOI Outcome 2 (IOI 2)'],x['Outcome1 Construct'],x['IOI Secondary Outcome 2 (IOI 2)'],x['Outcome2 Construct'],x['HRA Level of Service'],x['Date Closed'],x['IOI Outcome 2 Date (IOI 2)'],x['IOI Secondary Outcome Date 2 (IOI 2)']), axis=1)
        

        #kind of glitchy - if it has an outcome date but no outcome it doesn't say *Needs Outcome*
        
        #add LSNYC to start of case numbers 
        
        df['Unique_ID'] = 'LSNYC'+df['Matter/Case ID#']
        
        #take second letters of first and last names
        
        df['Last_Initial'] = df['Client Last Name'].str[1]
        df['First_Initial'] = df['Client First Name'].str[1]

        #Year of birth
        df['Year_of_Birth'] = df['Date of Birth'].str[-4:]
        
        #Unique Client ID#
        df['Unique Client ID#'] = df['First_Initial'] + df['Last_Initial'] + df['Year_of_Birth']       
        
        #Deliverable Categories
        
        def Deliverable_Category(HRA_Coded_Case,Income_Cleanup,Age_at_Intake,ClientsNames):
            if Income_Cleanup == 'Needs Income Waiver':
                return 'Needs Cleanup'
            elif HRA_Coded_Case == 'T2-RD' and Age_at_Intake <= 21:
                return 'Tier 2 (minor removal)'
            elif HRA_Coded_Case == 'T2-RD' and ClientsNames in ImmigrationToolBox.AtlasClientsNames :
                return 'Tier 2 (minor removal)'
            elif HRA_Coded_Case == 'T2-RD':
                return 'Tier 2 (removal)'
            elif HRA_Coded_Case.startswith('T2') == True:
                return 'Tier 2 (other)'
            elif HRA_Coded_Case.startswith('T1')== True:
                return 'Tier 1'
            elif HRA_Coded_Case.startswith('B') == True:
                return 'Brief'
            else:
                return 'Needs Cleanup'

        df['Deliverable Tally'] = df.apply(lambda x: Deliverable_Category(x['HRA Case Coding'],x['Exclude due to Income?'],x['Age at Intake'], x['Client Name']), axis=1)
        
        #make all cases for any client that has a minor removal tally, into also being minor removal cases
        
        
        df['Modified Deliverable Tally'] = ''
        
        dfs = df.groupby('Unique Client ID#',sort = False)

        tdf = pd.DataFrame()
        for x, y in dfs:
            for z in y['Deliverable Tally']:
                if z == 'Tier 2 (minor removal)':
                    y['Modified Deliverable Tally'] = 'Tier 2 (minor removal)'
            tdf = tdf.append(y)
        df = tdf
        
        
        #write function to identify blank 'modified deliverable tallies' and add it back in as the original deliverable tally
        df.fillna('',inplace= True)

        def fillBlanks(ModifiedTally,Tally):
            if ModifiedTally == '':
                return Tally
            else:
                return ModifiedTally
        df['Modified Deliverable Tally'] = df.apply(lambda x: fillBlanks(x['Modified Deliverable Tally'],x['Deliverable Tally']),axis=1)

        #Reportable?
        
        df['Reportable?'] = df.apply(lambda x: ImmigrationToolBox.ReportableTester(x['Exclude due to Income?'],x['Needs DHCI?'],x['Needs Substantial Activity?'],x['Deliverable Tally']),axis=1)
        
        #***add code to make it so that it deletes any extra 'brief' cases for clients that have mutliple cases
        
        
        
        
        
        #gender
        def HRAGender (gender):
            if gender == 'Male' or gender == 'Female':
                return gender
            else:
                return 'Other'
        df['Gender'] = df.apply(lambda x: HRAGender(x['Gender']), axis=1)
        
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
        df['Waiver_Type'] = df.apply(lambda x: IncWaiver(x['Income_Eligible'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)']), axis=1)
        
        def IncWaiverDate (waivertype,date):
            if waivertype != '':
                return date
            else:  
                return ''
                
        df['Waiver_Approval_Date'] = df.apply(lambda x: IncWaiverDate(x['Waiver_Type'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)']), axis = 1)
        
        #Referrals
        def Referral (referral):
            if referral == "Action NY":
                return "ActionNYC"
            elif referral == "HRA":
                return "HRA-DSS"
            elif referral == "Other":
                return "Other"
            elif referral == "":
                return "None"
            else:
                return ""
                
        df['Referral_Source'] = df.apply(lambda x: Referral(x['IOI Referral Source (IOI 2)']), axis = 1)
        
        #Pro Bono Involvement
        def ProBonoCase (branch, pai):
            if branch == "LSU" or pai == "Yes":
                return "YES"
            else:
                return "NO"
                
        df['Pro_Bono'] = df.apply(lambda x:ProBonoCase(x['Assigned Branch/CC'], x['PAI Case?']), axis = 1)
        
        #Prior Enrollment
        
        def PriorEnrollment (casenumber):
            if casenumber in ImmigrationToolBox.ReportedFY21:
                return 'FY 21'
            elif casenumber in ImmigrationToolBox.ReportedFY20:
                return 'FY 20'
            elif casenumber in ImmigrationToolBox.ReportedFY19:
                return 'FY 19'
                
                
        df['Prior_Enrollment_FY'] = df.apply(lambda x:PriorEnrollment(x['Matter/Case ID#']), axis = 1)
        
        #Other Cleanup
        df['Service_Type_Code'] = df['HRA Service Type']
        df['Proceeding_Type_Code'] = df['HRA Proceeding Type']
        df['Outcome'] = df['Outcome To Report']
        df['Outcome_Date'] = df['Outcome Date To Report']
        df['Seized_at_Border'] = df['IOI Was client apprehended at border? (IOI 2&3)']
        df['Group'] = ''
        
          
                
        #REPORTING VERSION Put everything in the right order
        df = df[['Unique_ID','Last_Initial','First_Initial','Year_of_Birth','Gender','Country of Origin','Borough','Zip Code','Language','Household_Size','Number_of_Children','Annual_Income','Income_Eligible','Waiver_Type','Waiver_Approval_Date','Eligibility_Date','Referral_Source','Service_Type_Code','Proceeding_Type_Code','Outcome','Outcome_Date','Seized_at_Border','Group','Prior_Enrollment_FY','Pro_Bono','Special Legal Problem Code','HRA Level of Service','HRA Case Coding','Hyperlinked Case #','Office','Primary Advocate','Client Name','Special Legal Problem Code','Level of Service','Needs DHCI?','Exclude due to Income?','Needs Substantial Activity?','IOI FY22 Substantial Activity 2022','IOI Date Substantial Activity Performed 2022','Country of Origin','Outcome To Report','HRA Case Coding','IOI Was client apprehended at border? (IOI 2&3)','Deliverable Tally','Modified Deliverable Tally','Reportable?']]
            
        
                
        
        #Preparing Excel Document
        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        problem_format = workbook.add_format({'bg_color':'yellow'})
        
        
        
        worksheet.set_column('B:B',19)
        worksheet.set_column('C:BL',30)
        
        EFRowRange='E1:F'+str(df.shape[0]+1)
        print("EFRowRange = "+EFRowRange)
        FRowRange='F1:F'+str(df.shape[0]+1)
        print("FRowRange = "+FRowRange)
        JKRowRange='J1:K'+str(df.shape[0]+1)
        print("JKRowRange = "+JKRowRange)
        
        worksheet.conditional_format(EFRowRange,{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
        worksheet.conditional_format(FRowRange,{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Hold for Review"',
                                                 'format': problem_format})
        worksheet.conditional_format('G1:G100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs DHCI Form"',
                                                 'format': problem_format})                                 
        worksheet.conditional_format('H1:H100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Income Waiver"',
                                                 'format': problem_format})
        worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Substantial Activity in FY21"',
                                                 'format': problem_format})
        worksheet.conditional_format(JKRowRange,{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
        worksheet.conditional_format('L1:L100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"*Needs Outcome*"',
                                                 'format': problem_format})                                         
        worksheet.conditional_format('L1:L100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"*Needs Outcome Date*"',
                                                 'format': problem_format})
        
        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Formatted " + f.filename)

    return '''
    <!doctype html>
    <title>IOI Immigration Quarterly</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Quarterly Formatting for IOI Immigration:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=IOI-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1918" target="_blank">"Grants Management IOI 2 (3459) Report"</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for IOI.</li> 
    <li>Once you have identified this file, click ‘IOI-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
