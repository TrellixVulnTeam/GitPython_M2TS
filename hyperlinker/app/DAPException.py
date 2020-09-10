#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os

@app.route("/DAPException", methods=['GET', 'POST'])
def DAPException():
    #upload file from computer
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        test = pd.read_excel(f)        
        test.fillna('',inplace=True)


        if test.iloc[0][0] == '':
            df = pd.read_excel(f,skiprows=2)
        else:
            df = pd.read_excel(f)
        df.fillna('',inplace=True)
        
        def NoIDDelete(CaseID):
            if CaseID == '':
                return 'No Case ID'
            else:
                return str(CaseID)
        df['Matter/Case ID#'] = df.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        
        df = df[df['Matter/Case ID#'] != 'No Case ID']

        if 'Matter/Case ID#' not in df.columns:
            df['Matter/Case ID#'] = df['id']

        last7 = df['Matter/Case ID#'].apply(lambda x: x[-7:])
        df['Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + df['Matter/Case ID#'] +'"' +')'
        move = df['Hyperlinked Case #']
        del df['Hyperlinked Case #']
        df.insert(0,'Hyperlinked Case #',move)           
        
        #Test if Social Security Number is Correctly Formatted
        
        def SSNum (CaseNum):
            CaseNum = str(CaseNum)
            First3 = CaseNum[0:3]
            Middle2 = CaseNum[4:6]
            Last4 = CaseNum[7:11]
            FirstDash = CaseNum[3:4]
            SecondDash = CaseNum[6:7]
                        
            if First3 == '000' and Middle2 == '00':
                return 'Needs  Full SS#'
            elif str.isnumeric(First3) == True and str.isnumeric(Middle2) == True and str.isnumeric(Last4) == True and FirstDash == '-' and SecondDash == '-': 
                return ''
            else:
                return "Needs Correct SS # Format"
                
        df['SS # Tester'] = df.apply(lambda x: SSNum(x['S.S.N.']), axis=1)
        
        #Is there a DAP Income Type for every Case?
        
        def DAPIncomeType (IncomeType):
            if IncomeType == '':
                return 'Needs DAP Income Type'
            else:
                return ''
        df['DAP Income Type Tester'] = df.apply(lambda x: DAPIncomeType(x['DAP Income Type']), axis=1)
        
        #Test for blank 'DAP Problem Type'
        
        def DAPProblemTester (DAPProblem):
            if DAPProblem == '':
                return 'Needs DAP Legal Problem'
            else:
                return ''
                
        df ['DAP Legal Problem Tester'] = df.apply(lambda x : DAPProblemTester(x['DAP Legal Problem']), axis=1)
        
        #Test for blank highest level of representation
        
        def DAPRepresentationTester (DAPRepresentation):
            if DAPRepresentation == '':
                return 'Needs Level of Representation'
            else:
                return ''
                
        df ['DAP Level of Representation Tester'] = df.apply(lambda x : DAPRepresentationTester(x['DAP Level Of Representation']), axis=1)
        
        #Test if there's an ALJ Name for cases that have ALJ Hearings
        
        def ALJNameTester (RepresentationLevel, ALJName, DAPOutcome):
            if RepresentationLevel == 'ALJ Hearing' and ALJName == '' and DAPOutcome != 'Short or other services':
                return 'Needs ALJ Name'
            else:
                return ''
        df['DAP ALJ Name Tester'] = df.apply(lambda x : ALJNameTester(x['DAP Level Of Representation'],x['DAP ALJ Name'],x['DAP Outcome']), axis = 1)
        
        
        #Convert LegalServer Names into Comprehensible Column Headers (only necessary if using old DAP legalserver report)
        
        #df['Received DIB?'] = df['Custom - DAP Disability - Title II']
        #df['Received SSI?'] = df['Custom - DAP SSI Title XVI']
        #df['Monthly DIB Award'] = df['Custom - DAP Monthly Disability']
        #df['Monthly SSI Award'] = df['Custom - DAP Monthly Social Security']
        #df['Retroactive Award'] = df['Custom - DAP Retro Total']
        #df['Interim Assistance'] = df['Custom - DAP Interim Deduction']
        
        
        #Test Monthly Award Amounts over 3k?
        
        def MonthlyAwardTester (SSIMonthly, DIBMonthly):
            if SSIMonthly >= 3000 or DIBMonthly >= 3000:
                return 'Needs to Confirm Monthly $ Amount'
            else:
                return ''
        df ['Monthly Award Tester'] = df.apply(lambda x : MonthlyAwardTester(x['Monthly SSI Award'],x['Monthly DIB Award']), axis = 1)        
        
        #Retro awards testing if they're over $100k
        
        def RetroAwardTester (DAPRetro):
            if DAPRetro >= 100000:
                return 'Needs to Confirm DAP Retro $ Amount'
            else:
                return ''
        df ['Retro Award Tester'] = df.apply(lambda x : RetroAwardTester(x['DAP Retroactive Award']), axis = 1)          
                
       
        
        #BlankOutcomeTester
                                
        def BlankOutcomeTester (DAPOutcome):
            if DAPOutcome == '':
               return 'Needs Outcome'
            else :
               return ''
        df ['Blank Outcome Tester'] = df.apply(lambda x : BlankOutcomeTester(x['DAP Outcome']), axis = 1)
        
        
        
        #If case has benefits, then it can't have short service as an outcome
        
        NoBenefitsList = ['Short or other services','Client did not receive / retain benefits','Client withdrew/failed to return','Client won/did not receive any benefits','Case Remanded']
        
        def DAPOutcomeTester (DAPOutcome,DAPRetro,DAPInterim,SSIMonthly,DIBMonthly,NoBenefitsList,LevelOfRep):
            if DAPOutcome in NoBenefitsList and DAPRetro > 0:
                return 'Needs Higher Level Outcome'
            elif DAPOutcome in NoBenefitsList and DAPInterim > 0:
                return 'Needs Higher Level Outcome'           
            elif DAPOutcome in NoBenefitsList and SSIMonthly > 0:
                return 'Needs Higher Level Outcome'
            elif DAPOutcome in NoBenefitsList and DIBMonthly > 0:
                return 'Needs Higher Level Outcome'
            elif LevelOfRep == 'ALJ Hearing' and DAPOutcome == 'Short or other services':
                return 'Needs Higher Level Outcome'
            else :
                return ''
        df ['DAP Outcome Tester'] = df.apply(lambda x : DAPOutcomeTester(x['DAP Outcome'],x['DAP Retroactive Award'],x['Interim Assistance'],x['Monthly SSI Award'],x['Monthly DIB Award'],NoBenefitsList,x['DAP Level Of Representation']), axis = 1)
        
        
        #if the outcome is monthly benefits then either DAP Monthly or SSI Monthly has to have $
        def MonthlyBenefitsTester  (DAPOutcome,DIBMonthly,SSIMonthly):
            DIBMonthly = int(DIBMonthly)
            SSIMonthly = int(SSIMonthly)
            
            if DAPOutcome == 'Client won/ received  retained monthly benefits' and DIBMonthly == 0 and SSIMonthly == 0:
                return 'Needs Monthly Benefits'
            else:
                return ''
        df ['DAP Monthly Benefits Tester'] = df.apply(lambda x : MonthlyBenefitsTester(x['DAP Outcome'],x['Monthly DIB Award'],x['Monthly SSI Award']),axis = 1)
       
       
        #If the outcome says you got retro, must enter retro award $
       
        def DAPRetroTester (DAPOutcome,DAPRetro):
            DAPRetro = int(DAPRetro)
            if DAPOutcome == 'Client won/received only retroactive benefits' and DAPRetro == 0:
                return "Needs Retro Award $"
            else:
                return ''
        df ['DAP Retro Tester'] = df.apply(lambda x : DAPRetroTester(x['DAP Outcome'],x['DAP Retroactive Award']),axis = 1)
        
        
        #if outcome = no benefits, then there should not be $ benefits
        
        def NoBenefitsTester (DAPOutcome,DAPRetro,DAPInterim,SSIMonthly,DIBMonthly,NoBenefitsList):
            if DAPOutcome in NoBenefitsList and DAPRetro > 0:
                return 'Should Not have $ Benefits with this Outcome'
            elif DAPOutcome in NoBenefitsList and DAPInterim > 0:
                return 'Should Not have $ Benefits with this Outcome'
            elif DAPOutcome in NoBenefitsList and SSIMonthly > 0:
                return 'Should Not have $ Benefits with this Outcome'
            elif DAPOutcome in NoBenefitsList and DIBMonthly > 0:
                return 'Should Not have $ Benefits with this Outcome'
            else :
                return ''
        df ['No Benefits Tester'] = df.apply(lambda x : NoBenefitsTester(x['DAP Outcome'],x['DAP Retroactive Award'],x['Interim Assistance'],x['Monthly SSI Award'],x['Monthly DIB Award'],NoBenefitsList), axis = 1)
        
        #Received SSI Tester
        
        def SSITester (ReceivedSSI,SSIMonthly,ReceivedDIB,DAPOutcome):
            SSIMonthly = int(SSIMonthly)
            if ReceivedSSI == 'Yes' and SSIMonthly == 0:
                return 'Needs SSI Award Amount'
            elif ReceivedSSI != 'Yes' and SSIMonthly > 0:
                return 'Needs Received SSI? = Yes'
            elif DAPOutcome == 'Client won/received only retroactive benefits' and ReceivedDIB != 'Yes' and ReceivedSSI != 'Yes':
                return 'Needs Received DIB or SSI = Yes'
            else:
                return ''
                
        df ['SSI Tester'] = df.apply(lambda x : SSITester(x['Received SSI?'],x['Monthly SSI Award'],x['Received DIB?'],x['DAP Outcome']),axis = 1)
                
        
        #Received DIB Tester
        
        def DIBTester (ReceivedDIB,DIBMonthly,ReceivedSSI,DAPOutcome):
            DIBMonthly = int(DIBMonthly)
            if ReceivedDIB == 'Yes' and DIBMonthly == 0:
                return 'Needs DIB Award Amount'
            elif ReceivedDIB != 'Yes' and DIBMonthly > 0:
                return 'Needs Received DIB? = Yes'
            elif DAPOutcome == 'Client won/received only retroactive benefits' and ReceivedDIB != 'Yes' and ReceivedSSI != 'Yes':
                return 'Needs Received DIB or SSI = Yes'
            else:
                return ''
                
        df ['DIB Tester'] = df.apply(lambda x : DIBTester(x['Received DIB?'],x['Monthly DIB Award'],x['Received SSI?'],x['DAP Outcome']),axis = 1)
        
        #If Client withdrew - Do Not Report
        
        def WithdrewTester(DAPOutcome):
            if DAPOutcome == 'Client withdrew/failed to return':
                return 'Client Withdrew - Should Not Report'
            else:
                return ''
        
        df ['Withdrew Tester'] = df.apply(lambda x : WithdrewTester(x['DAP Outcome']),axis = 1)
        
        #Tester Tester
        
        def TesterTester (SSTester,IncomeTypeTester,LegalProblemTester,LevelofRepTester,ALJNameTester,MonthlyAwardTester,RetroAwardTester,BlankOutcomeTester,OutcomeTester,MonthlyBenefitsTester,RetroBenefitsTester,NoBenefitsTester,SSITester,DIBTester,WithdrewTester):
            if SSTester == '' and IncomeTypeTester == "" and LegalProblemTester  == "" and LevelofRepTester == "" and ALJNameTester == "" and MonthlyAwardTester == "" and RetroAwardTester == "" and BlankOutcomeTester == '' and OutcomeTester == "" and MonthlyBenefitsTester == "" and RetroBenefitsTester == "" and NoBenefitsTester == '' and SSITester == '' and DIBTester == '' and WithdrewTester == '':
                return 'Prepared for Submission'
            else:
                return 'Case Needs Attention'
                
        df ['Case Attention Tester'] = df.apply(lambda x: TesterTester(x['SS # Tester'],x['DAP Income Type Tester'],x['DAP Legal Problem Tester'],x['DAP Level of Representation Tester'],x['DAP ALJ Name Tester'],x['Monthly Award Tester'],x['Retro Award Tester'],x['Blank Outcome Tester'],x['DAP Outcome Tester'],x['DAP Monthly Benefits Tester'],x['DAP Retro Tester'],x['No Benefits Tester'],x['SSI Tester'],x['DIB Tester'],x['Withdrew Tester']), axis = 1)
        
        #Columns for Formatted Report
        
        df['Agency'] = 'LS-NYC'
        df['Region'] = 'NYC'
        df['Agency Office'] = '0'
        
        
        def CaseSplicer (CaseID):
            FirstChar = CaseID[0:1]
            if str.isalpha(FirstChar) == True:
                return CaseID[1:3] + CaseID[7:]
            else:
                return CaseID[0:2] + CaseID[4:]
        df['Case#'] = df.apply(lambda x: CaseSplicer(x['Matter/Case ID#']), axis =1)
        
        df['Case open date'] = df['Date Opened']
        df['Case close date'] = df['Date Closed']
        df['Client name'] = df['Client Name']
        df['Client SS#'] = df['S.S.N.']
        df['Client DOB'] = df['Date of Birth']
        df['Client gender'] = df['Gender']
        df['Client ethnicity'] = df['Race']
        df['Client county'] = df['County of Residence']
        df['Client ZIP code'] = df['Zip Code']
        df['Client disabilities'] = ''
        
        #if dap income is DAP eligible w/out PA
        
        def EligforPA(DAPIncome):
            if DAPIncome == "DAP eligible w/o PA":
                return "T"
            else: 
                return "F"
        
        df['Elig for PA w/o SSI/SSD'] = df.apply(lambda x : EligforPA(x['DAP Income Type']),axis = 1)
      
      #is it an old funding code 2070
      
        def DAPTANF(FundingCode):
            if FundingCode.startswith("2070") == True:
                return "T"
            else: 
                return "F"
        
        df['DAP TANF'] = df.apply(lambda x : DAPTANF(x['Primary Funding Codes']),axis = 1)
        
        def PATranslator(DAPIncomeType):
            if DAPIncomeType == "TANF":
                return "1"
            elif DAPIncomeType == "SN or HASA or PA":
                return "2"
            elif DAPIncomeType == "Medicaid":
                return "3"
            elif DAPIncomeType == "DAP eligible w/o PA" or DAPIncomeType == "Eligible for but not receiving PA":
                return "4"
        df['PA category'] = df.apply(lambda x : PATranslator(x['DAP Income Type']),axis=1)
        
        def ReferralTranslator(ReferralType):
            if ReferralType == "Former Client":
                return "1"
            elif ReferralType == "Friend":
                return "2"
            elif ReferralType == "Local DSS":
                return "3"
            elif ReferralType == "SSA":
                return "4"
            else:
                return "5"
        df['Referral source'] = df.apply(lambda x : ReferralTranslator(x['DAP Referral Source']),axis=1)

        
        #Local DSS placement
      
        def LocalDSS(ReferralSource):
            if ReferralSource == "Local DSS":
                return "HRA"
            else: 
                return ""
        
        df['DSS region'] = df.apply(lambda x : LocalDSS(x['DAP Referral Source']),axis = 1)
       
        def ProblemTranslator(DAPLegalProblem):
            if DAPLegalProblem == "No Problem":
                return "1"
            elif DAPLegalProblem == "Claim Denial":
                return "2"
            elif DAPLegalProblem == "Termination":
                return "3"
            elif DAPLegalProblem == "Other":
                return "4"
        df['SSI/SSD problem'] = df.apply(lambda x : ProblemTranslator(x['DAP Legal Problem']),axis=1)
        
        
        def LevelTranslator(Level):
            if Level == "Reconsideration'":
                return "A"
            elif Level == "ALJ Hearing":
                return "B"
            elif Level == "Appeals Council":
                return "C"
            elif Level == "District Court":
                return "D"
            elif Level == "Other":
                return "E"
        df['Highest level of review'] = df.apply(lambda x : LevelTranslator(x['DAP Level Of Representation']),axis=1)
        
        df['ALJ name'] = df['DAP ALJ Name']
        
        def OutcomeTranslator(DAPOutcome):
            if DAPOutcome == "Client did not receive / retain benefits":
                return "1"
            elif DAPOutcome == "Client withdrew/failed to return":
                return "2"
            elif DAPOutcome == "Case Remanded":
                return "3"
            elif DAPOutcome == "Short or other services":
                return "4"
            elif DAPOutcome == "Client won/ received  retained monthly benefits":
                return "5"
            elif DAPOutcome == "Client won/received only retroactive benefits":
                return "6"
            elif DAPOutcome == "Client won/did not receive any benefits":
                return "7"
        df['Outcome'] = df.apply(lambda x : OutcomeTranslator(x['DAP Outcome']),axis=1)
        
        
        #Swap Yes for T and No for F
        def ReceivedDIB(ReceivedDIB):
            if ReceivedDIB == "Yes":
                return "T"
            elif ReceivedDIB == "No":
                return "F"
            else: 
                return "F"
        df['Received DIB'] = df.apply(lambda x : ReceivedDIB(x['Received DIB?']),axis = 1)
        
        def ReceivedSSI(ReceivedSSI):
            if ReceivedSSI == "Yes":
                return "T"
            elif ReceivedSSI == "No":
                return "F"
            else: 
                return "F"
        df['Received SSI'] = df.apply(lambda x : ReceivedSSI(x['Received SSI?']),axis = 1)
        
        
        df['Monthly DIB Award'] = df['Monthly DIB Award'].replace('0',0)
        df['Monthly SSI Award'] = df['Monthly SSI Award'].replace('0',0)
        df['DAP Retroactive Award'] = df['DAP Retroactive Award'].replace('0',0)
        df['Interim Assistance'] = df['Interim Assistance'].replace('0',0)
        
        
        df['DIB award amount'] = round(df['Monthly DIB Award'])
        df['SSI award amount'] = round(df['Monthly SSI Award'])
        df['Retroactive award amt'] = round(df['DAP Retroactive Award'])
        df['Interim assistance amt'] = round(df['Interim Assistance'])
        #***these dollar values need to be rounded to nearest dollar
        
        
        
        #Ordering Spreadsheet Correctly
        
        df = df[['Hyperlinked Case #','Assigned Branch/CC','Primary Advocate','Client Name',
        'Case Attention Tester',
        'Withdrew Tester',
        'S.S.N.','SS # Tester',
        'DAP Income Type','DAP Income Type Tester',
        'DAP Legal Problem', 'DAP Legal Problem Tester',
        'DAP Level Of Representation','DAP Level of Representation Tester',
        'DAP ALJ Name','DAP ALJ Name Tester',
        'Monthly SSI Award','Monthly DIB Award','Monthly Award Tester',
        'DAP Retroactive Award','Retro Award Tester',
        'DAP Outcome','Blank Outcome Tester','DAP Outcome Tester',
        'DAP Monthly Benefits Tester',
        'DAP Retro Tester',
        'No Benefits Tester',
        'Received SSI?','SSI Tester',
        'Received DIB?','DIB Tester',
        'Agency',
        'Region',
        'Agency Office',
        'Case#',
        'Case open date',
        'Case close date',
        'Client name',
        'Client SS#',
        'Client DOB',
        'Client gender',
        'Client ethnicity',
        'Client county',
        'Client ZIP code',
        'Client disabilities',
        'Elig for PA w/o SSI/SSD',
        'DAP TANF',
        'PA category',
        'Referral source',
        'DSS region',
        'SSI/SSD problem',
        'Highest level of review',
        'ALJ name',
        'Outcome',
        'Received DIB',
        'Received SSI',
        'DIB award amount',
        'SSI award amount',
        'Retroactive award amt',
        'Interim assistance amt'
       
        ]]       

        allgood_dictionary = dict(tuple(df.groupby('Case Attention Tester')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                worksheet = writer.sheets[i]
                worksheet.freeze_panes(1,1)
                worksheet.autofilter('A1:ZZ1')
                #create format that will make case #s look like links
                link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
       
                #assign new format to column A
                worksheet.set_column('A:A',20,link_format)
        
                worksheet.set_column('B:ZZ',30)
        
                worksheet.conditional_format('B1:ZZ1000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
                worksheet.conditional_format('B1:ZZ1000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Should',
                                                 'format': problem_format})
                
            writer.save()
        
        
        output_filename = f.filename
        
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename)
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "DAPCleaned " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>DAP Exception Report</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>DAP Exceptions & Report Prep Tool</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Cleanup!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2288" target="_blank">"DAP Exceptions for Python Tool"</a>.</li>
    <li>This tool will create a spreadsheet with 2 tabs, the first "Case Needs Attention" should be used to help advocates identify data gaps in their cases. The second "Prepared for Submission" contains re-formatted case data in columns AF-BH that are suitable for submission to EJC.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for error-checking & report prep.</li> 
    <li>Once you have identified this file, click ‘Cleanup!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
