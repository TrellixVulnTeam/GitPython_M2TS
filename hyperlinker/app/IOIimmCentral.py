from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db , ImmigrationToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd



@app.route("/IOIimmCentral", methods=['GET', 'POST'])
def upload_IOIimmCentral():
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
        move=data_xls['Temp Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Bronx Legal Services','BxLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Brooklyn Legal Services','BkLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Queens Legal Services','QLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Manhattan Legal Services','MLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Staten Island Legal Services','SILS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Legal Support Unit','LSU')
        
        
        
        #Determining 'level of service' from 3 fields       
        def HRA_Service_Level(Close_Reason,LS_LoS):
            Close_Reason = str(Close_Reason)
            LS_LoS = str(LS_LoS)
                
            if Close_Reason.startswith("A") == True:
                return 'Advice'
            elif Close_Reason.startswith("B") == True:
                return 'Brief Service'
            elif Close_Reason.startswith("H") == True or Close_Reason.startswith("I") == True or Close_Reason.startswith("L"):
                return 'Full Rep or Extensive Service'
            elif LS_LoS == 'Advice':
                return 'Advice'
            elif LS_LoS == 'Hold For Review':
                return 'Hold For Review'
            elif LS_LoS == 'Brief Service' or LS_LoS == 'Out-of-Court Advocacy':
                return 'Brief Service'
            elif LS_LoS.startswith('Rep') == True:
                return 'Full Rep or Extensive Service'
            else:
                return ''
        
        data_xls['HRA Level of Service'] = data_xls.apply(lambda x: HRA_Service_Level(x['Close Reason'],x['Level of Service']), axis=1)
        
        
        #Putting Cases into HRA's Baskets!
        def HRA_Case_Coding(LPC,SLPC,HRA_LoS,Crim):
            LPC = str(LPC)
            SLPC = str(SLPC)
            if HRA_LoS == '***Needs Cleanup***':
                return ''
            elif HRA_LoS == 'Hold For Review':
                return 'Hold For Review'
            elif LPC.startswith('2') == True and HRA_LoS == 'Advice':
                return 'B -EMP'
            elif LPC.startswith('2') == True and HRA_LoS == 'Brief Service':
                return 'B -EMP'
            elif SLPC == 'G-639' and HRA_LoS == 'Advice':
                return 'B -INQ'
            elif SLPC == 'G-639' and HRA_LoS == 'Brief Service':
                return 'B -INQ'
            elif SLPC == 'I-914' and HRA_LoS == 'Advice':
                return 'B -CERT'
            elif SLPC == 'I-914' and HRA_LoS == 'Brief Service':
                return 'B -CERT'
            elif SLPC == 'I-918' and HRA_LoS == 'Advice':
                return 'B -CERT'
            elif SLPC == 'I-918' and HRA_LoS == 'Brief Service':
                return 'B -CERT'
            elif HRA_LoS == 'Advice' or HRA_LoS == 'Brief Service':
                return 'B -ADVI'
            elif SLPC == "Emergency Planning":
                return 'B -APD'    
            elif SLPC == "Parental Designation Form":
                return 'B -OTH_Parental Designation Form'
            elif Crim == "Yes":
                return 'T2-OTH_CRM'
            elif SLPC == "I-589 Affirmative" or SLPC == "I-730":
                return 'T2-AR'
            elif SLPC == "I-589 Defensive" or SLPC == "Removal Defense" or SLPC == "EOIR-40" or SLPC == "EOIR-42A"or SLPC == "EOIR-42B" or SLPC == "I-212" or SLPC == "I-485 Defensive":
                return 'T2-RD'
            elif SLPC == "I-912":
                return 'T1-OTH_I-912'
            elif SLPC == "I-130 (spouse)":
                return 'T2-MAR'
            elif SLPC == "I-129F" or SLPC == "I-130" or SLPC == "I-751" or SLPC == "I-864"or SLPC == "I-864EZ" or SLPC == "AOS I-130":
                return 'T1-FAM'
            elif SLPC == "204(L)":
                return 'T2-HO_204(L)'
            elif SLPC == "AR-11":
                return 'T1-OTH_AR-11'    
            elif SLPC == "DS-160" or SLPC == "DS-260":
                return 'T1-CON'
            elif SLPC == "EOIR 27" or SLPC == "Mandamus Action" or SLPC == "EOIR-26":
                return 'T2-FED'    
            elif SLPC == "EOIR-29 BIA Appeal" or SLPC.startswith("I-290B") == True or SLPC == "N-336":
                return 'T2-APO'
            elif SLPC == "G-639":
                return 'T1-OTH_G639'
            elif SLPC == "I-102":
                return 'T1-OTH_I102'
            elif SLPC.startswith("I-131") == True:
                return 'T1-TRV'
            elif SLPC == "I-192":
                return 'T2-OTH_I-192'
            elif SLPC == "I-360 SIJS" or LPC == "44 Minor Guardianship / Conservatorship" or LPC == "42 Neglected/Abused/Dependent":
                return 'T2-SIJS'
            elif SLPC == "I-360 VAWA Self-Petition":
                return 'T2-VAWA'
            elif SLPC == "I-539":
                return 'T1-OTH_I539'
            elif SLPC == "I-601" or SLPC == "I-601A":
                return 'T2-WOI'
            elif SLPC == "I-765":
                return 'T1-EAD'
            elif SLPC == "I-821":
                return 'T1-TPS'
            elif SLPC == "I-821D":
                return 'T1-DACA'
            elif SLPC == "I-824":
                return 'T1-OTH_I824'
            elif SLPC == "I-864W":
                return 'T2-VAWA'
            elif SLPC == "I-881":
                return 'T2-HO_I881'
            elif SLPC == "I-90":
                return 'T1-GCR'
            elif SLPC == "I-914" or SLPC == "I-914A":
                return 'T2-HO_I914'
            elif SLPC == "I-918" or SLPC == "I-918A":
                return 'T2-HO_I918'
            elif SLPC == "I-929":
                return 'T2-HO_I929'
            elif SLPC == "I-942":
                return 'T1-OTH_I942'
            elif SLPC == "N-400" or SLPC == "N-565" or SLPC == "N-600" or SLPC == "N-648":
                return 'T1-CIT'
            elif SLPC == "327 Uncontested Divorce":
                return 'T1-SC'
            elif SLPC == "Contested Divorce":
                return 'T2-SC'
            elif SLPC == "340 Name Change":
                return 'T1-OTH_NameChange'
            elif SLPC == "311 Custody":
                return 'T2-FC'
            elif SLPC == 'Legal Rep - Employment Law':
                return 'T1-EMP'
            elif SLPC == 'AOS (Not I-130)' or SLPC == "I-485 Affirmative" :
                return 'T1-AOS'
            elif SLPC == "EOIR-26A":
                return 'T2-OTH_EOIR26a'
            elif SLPC == "EOIR-33/BIA" or SLPC == 'EOIR-33/IC':
                return 'T2-OTH_EOIR33'
            elif SLPC == "I-134":
                return 'T2-OTH_I134'
            else:
                return 'Something is wrong'
    
        data_xls['HRA Case Coding'] = data_xls.apply(lambda x: HRA_Case_Coding(x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA Level of Service'],x['IOI Does Client Have A Criminal History? (IOI 2)']), axis=1)

        #Dummy SLPC for Juvenile Cases
        def DummySLPC(LPC,SLPC):
            LPC = str(LPC)
            SLPC = str(SLPC)
            if LPC == "44 Minor Guardianship / Conservatorship" or LPC == "42 Neglected/Abused/Dependent":
                return 'N/A'
            else:
                return SLPC
                
        data_xls['Special Legal Problem Code'] = data_xls.apply(lambda x: DummySLPC(x['Legal Problem Code'],x['Special Legal Problem Code']), axis=1)
    
        data_xls['HRA Service Type'] = data_xls['HRA Case Coding'].apply(lambda x: x[:2] if x != 'Hold For Review' else '')

        data_xls['HRA Proceeding Type'] = data_xls['HRA Case Coding'].apply(lambda x: x[3:] if x != 'Hold For Review' else '')
        
        #Giving things better names in cleanup sheet
        
        data_xls['DHCI form?'] = data_xls['Has Declaration of Household Composition and Income (DHCI) Form?']
        
        data_xls['Consent form?'] = data_xls['IOI HRA Consent Form? (IOI 2)']
        
        data_xls['Client Name'] = data_xls['Full Person/Group Name (Last First)']
        
        data_xls['Office'] = data_xls['Assigned Branch/CC']
        
        data_xls['Country of Origin'] = data_xls['IOI Country Of Origin (IOI 1 and 2)']
        
        data_xls['Substantial Activity'] = data_xls['IOI Substantial Activity (Choose One)']
        
        data_xls['Date of Substantial Activity'] = data_xls['Custom IOI Date substantial Activity Performed']
        
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

        data_xls['Exclude due to Income?'] = data_xls.apply(lambda x: Income_Exclude(x['Percentage of Poverty'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)'],x['IOI Referral Source (IOI 2)']), axis=1)
        
        #Eligibility_Date & Rollovers 
        
        def Eligibility_Date(Effective_Date,Date_Opened):
            if Effective_Date != '':
                return Effective_Date
            else:
                return Date_Opened
        data_xls['Eligibility_Date'] = data_xls.apply(lambda x : Eligibility_Date(x['IOI HRA Effective Date (optional) (IOI 2)'],x['Date Opened']), axis = 1)
        
        #Manipulable Dates               
        
        data_xls['Open Month'] = data_xls['Eligibility_Date'].apply(lambda x: str(x)[:2])
        data_xls['Open Day'] = data_xls['Eligibility_Date'].apply(lambda x: str(x)[3:5])
        data_xls['Open Year'] = data_xls['Eligibility_Date'].apply(lambda x: str(x)[6:])
        data_xls['Open Construct'] = data_xls['Open Year'] + data_xls['Open Month'] + data_xls['Open Day']
        
        data_xls['Subs Month'] = data_xls['Date of Substantial Activity'].apply(lambda x: str(x)[:2])
        data_xls['Subs Day'] = data_xls['Date of Substantial Activity'].apply(lambda x: str(x)[3:5])
        data_xls['Subs Year'] = data_xls['Date of Substantial Activity'].apply(lambda x: str(x)[6:])
        data_xls['Subs Construct'] = data_xls['Subs Year'] + data_xls['Subs Month'] + data_xls['Subs Day']
        data_xls['Subs Construct'] = data_xls.apply(lambda x : x['Subs Construct'] if x['Subs Construct'] != '' else 0, axis = 1)
        
        data_xls['Outcome1 Month'] = data_xls['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[:2])
        data_xls['Outcome1 Day'] = data_xls['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[3:5])
        data_xls['Outcome1 Year'] = data_xls['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[6:])
        data_xls['Outcome1 Construct'] = data_xls['Outcome1 Year'] + data_xls['Outcome1 Month'] + data_xls['Outcome1 Day']

        data_xls['Outcome2 Month'] = data_xls['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[:2])
        data_xls['Outcome2 Day'] = data_xls['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[3:5])
        data_xls['Outcome2 Year'] = data_xls['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[6:])
        data_xls['Outcome2 Construct'] = data_xls['Outcome2 Year'] + data_xls['Outcome2 Month'] + data_xls['Outcome2 Day']       
        
        
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
        
        data_xls['Needs DHCI?'] = data_xls.apply(lambda x: DHCI_Needed(x['Has Declaration of Household Composition and Income (DHCI) Form?'],x['Open Construct'],x['Level of Service']), axis=1)
        
        
        #Needs Substantial Activity to Rollover into FY'20
        
        def Needs_Rollover(Open_Construct,Substantial_Activity,Substantial_Activity_Date,CaseID,ReportedFY19):
            if int(Open_Construct) >= 20190701:
                return ''
            elif Substantial_Activity != '' and int(Substantial_Activity_Date) >= 20190701 and int(Substantial_Activity_Date) <= 20200630:
                return ''
            elif CaseID in ReportedFY19:
                return 'Needs Substantial Activity in FY20'
            else:
                return ''
                
        data_xls['Needs Substantial Activity?'] = data_xls.apply(lambda x: Needs_Rollover(x['Open Construct'],x['Substantial Activity'],x['Subs Construct'],x['Matter/Case ID#'],x['ImmigrationToolBox.ReportedFY19']), axis=1)


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
        
        data_xls['Outcome To Report'] = data_xls.apply(lambda x: OutcomeToReport(x['IOI Outcome 2 (IOI 2)'],x['Outcome1 Construct'],x['IOI Secondary Outcome 2 (IOI 2)'],x['Outcome2 Construct'],x['HRA Level of Service'],x['Date Closed']), axis=1)
      
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
            
                
        data_xls['Outcome Date To Report'] = data_xls.apply(lambda x: OutcomeDateToReport(x['IOI Outcome 2 (IOI 2)'],x['Outcome1 Construct'],x['IOI Secondary Outcome 2 (IOI 2)'],x['Outcome2 Construct'],x['HRA Level of Service'],x['Date Closed'],x['IOI Outcome 2 Date (IOI 2)'],x['IOI Secondary Outcome Date 2 (IOI 2)']), axis=1)
        

        #kind of glitchy - if it has an outcome date but no outcome it doesn't say *Needs Outcome*
               
        
        #Deliverable Categories
        
        def Deliverable_Category(HRA_Coded_Case,Income_Cleanup,Age_at_Intake):
            if Income_Cleanup == 'Needs Income Waiver':
                return 'Needs Cleanup'
            elif HRA_Coded_Case == 'T2-RD' and Age_at_Intake <= 21:
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

        data_xls['Deliverable Tally'] = data_xls.apply(lambda x: Deliverable_Category(x['HRA Case Coding'],x['Exclude due to Income?'],x['Age at Intake']), axis=1)
        
        
        #add LSNYC to start of case numbers 
        
        data_xls['Unique_ID'] = 'LSNYC'+data_xls['Matter/Case ID#']
        
        #take second letters of first and last names
        
        data_xls['Last_Initial'] = data_xls['Client Last Name'].str[1]
        data_xls['First_Initial'] = data_xls['Client First Name'].str[1]

        #Year of birth
        data_xls['Year_of_Birth'] = data_xls['Date of Birth'].str[-4:]
        
        #gender
        def HRAGender (gender):
            if gender == 'Male' or gender == 'Female':
                return gender
            else:
                return 'Other'
        data_xls['Gender'] = data_xls.apply(lambda x: HRAGender(x['Gender']), axis=1)
        
        #county=borough
        data_xls['Borough'] = data_xls['County of Residence']
        
        #household size etc.
        data_xls['Household_Size'] = data_xls['Number of People under 18'].astype(int) + data_xls['Number of People 18 and Over'].astype(int)
        data_xls['Number_of_Children'] = data_xls['Number of People under 18']
        
        #Income Eligible?
        data_xls['Annual_Income'] = data_xls['Total Annual Income ']
        def HRAIncElig (PercentOfPoverty):
            if PercentOfPoverty > 200:
                return 'NO'
            else:
                return 'YES'
        data_xls['Income_Eligible'] = data_xls.apply(lambda x: HRAIncElig(x['Percentage of Poverty']), axis=1)
        
        def IncWaiver (eligible,waiverdate):
            if eligible == 'NO' and waiverdate != '':
                return 'Income'
            else:
                return ''
        data_xls['Waiver_Type'] = data_xls.apply(lambda x: IncWaiver(x['Income_Eligible'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)']), axis=1)
        
        def IncWaiverDate (waivertype,date):
            if waivertype != '':
                return date
            else:  
                return ''
                
        data_xls['Waiver_Approval_Date'] = data_xls.apply(lambda x: IncWaiverDate(x['Waiver_Type'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)']), axis = 1)
        
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
                
        data_xls['Referral_Source'] = data_xls.apply(lambda x: Referral(x['IOI Referral Source (IOI 2)']), axis = 1)
        
        #Pro Bono Involvement
        def ProBonoCase (branch, pai):
            if branch == "LSU" or pai == "Yes":
                return "YES"
            else:
                return "NO"
                
        data_xls['Pro_Bono'] = data_xls.apply(lambda x:ProBonoCase(x['Assigned Branch/CC'], x['PAI Case?']), axis = 1)
        
        
        #Other Cleanup
        data_xls['Service_Type_Code'] = data_xls['HRA Service Type']
        data_xls['Proceeding_Type_Code'] = data_xls['HRA Proceeding Type']
        data_xls['Outcome'] = data_xls['Outcome To Report']
        data_xls['Outcome_Date'] = data_xls['Outcome Date To Report']
        data_xls['Seized_at_Border'] = data_xls['IOI Was client apprehended at border? (IOI 2&3)']
        data_xls['Group'] = ''
        data_xls['Prior_Enrollment_FY'] = 'Jay does this manually later'
        data_xls = data_xls.sort_values(by=['Primary Advocate'])
          
        #CLEANUP VERSION Put everything in the right order
        data_xls = data_xls[['Hyperlinked Case #','Office','Primary Advocate','Client Name','Special Legal Problem Code','Level of Service','Needs DHCI?','Exclude due to Income?','Needs Substantial Activity?','Country of Origin','Language','Outcome To Report','IOI Was client apprehended at border? (IOI 2&3)','Deliverable Tally']]
        
        #Preparing Excel Document
        
        output_filename = f.filename

        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
            
        data_xls.to_excel(writer, sheet_name='Sheet1', index = False)
        workbook = writer.book
        link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column('A:A',20,link_format)
        worksheet.set_column('B:B',19)
        worksheet.set_column('C:BL',30)
                
        writer.save()
        
        
        
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>IOI Immigration for Central</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>IOI Immigration for Central:</h1>
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
