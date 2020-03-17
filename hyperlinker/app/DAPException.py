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
            data_xls = pd.read_excel(f,skiprows=2)
        else:
            data_xls = pd.read_excel(f)
        data_xls.fillna('',inplace=True)
        
        def NoIDDelete(CaseID):
            if CaseID == '':
                return 'No Case ID'
            else:
                return str(CaseID)
        data_xls['Matter/Case ID#'] = data_xls.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        
        data_xls = data_xls[data_xls['Matter/Case ID#'] != 'No Case ID']

        if 'Matter/Case ID#' not in data_xls.columns:
            data_xls['Matter/Case ID#'] = data_xls['id']

        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[-7:])
        data_xls['Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + data_xls['Matter/Case ID#'] +'"' +')'
        move = data_xls['Hyperlinked Case #']
        del data_xls['Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #',move)           
        del data_xls['Matter/Case ID#']
        
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
                
        data_xls['SS # Tester'] = data_xls.apply(lambda x: SSNum(x['S.S.N.']), axis=1)
        
        #Is there a DAP Income Type for every Case?
        
        def DAPIncomeType (IncomeType):
            if IncomeType == '':
                return 'Needs DAP Income Type'
            else:
                return ''
        data_xls['DAP Income Type Tester'] = data_xls.apply(lambda x: DAPIncomeType(x['DAP Income Type']), axis=1)
        
        #Test for blank 'DAP Problem Type'
        
        def DAPProblemTester (DAPProblem):
            if DAPProblem == '':
                return 'Needs DAP Legal Problem'
            else:
                return ''
                
        data_xls ['DAP Legal Problem Tester'] = data_xls.apply(lambda x : DAPProblemTester(x['DAP Legal Problem']), axis=1)
        
        #Test for blank highest level of representation
        
        def DAPRepresentationTester (DAPRepresentation):
            if DAPRepresentation == '':
                return 'Needs Level of Representation'
            else:
                return ''
                
        data_xls ['DAP Level of Representation Tester'] = data_xls.apply(lambda x : DAPRepresentationTester(x['DAP Level Of Representation']), axis=1)
        
        #Test if there's an ALJ Name for cases that have ALJ Hearings
        
        def ALJNameTester (RepresentationLevel, ALJName):
            if RepresentationLevel == 'ALJ Hearing' and ALJName == '':
                return 'Needs ALJ Name'
            else:
                ''
        data_xls['DAP ALJ Name Tester'] = data_xls.apply(lambda x : ALJNameTester(x['DAP Level Of Representation'],x['DAP ALJ Name']), axis = 1)
        
        
        #Convert LegalServer Names into Comprehensible Column Headers (only necessary if using old DAP legalserver report)
        
        #data_xls['Received DIB?'] = data_xls['Custom - DAP Disability - Title II']
        #data_xls['Received SSI?'] = data_xls['Custom - DAP SSI Title XVI']
        #data_xls['Monthly DIB Award'] = data_xls['Custom - DAP Monthly Disability']
        #data_xls['Monthly SSI Award'] = data_xls['Custom - DAP Monthly Social Security']
        #data_xls['Retroactive Award'] = data_xls['Custom - DAP Retro Total']
        #data_xls['Interim Assistance'] = data_xls['Custom - DAP Interim Deduction']
        
        
        #Test Monthly Award Amounts over 3k?
        
        def MonthlyAwardTester (SSIMonthly, DIBMonthly):
            if SSIMonthly >= 3000 or DIBMonthly >= 3000:
                return 'Needs to Confirm Monthly $ Amount'
            else:
                return ''
        data_xls ['Monthly Award Tester'] = data_xls.apply(lambda x : MonthlyAwardTester(x['Monthly SSI Award'],x['Monthly DIB Award']), axis = 1)        
        
        #Retro awards testing if they're over $100k
        
        def RetroAwardTester (DAPRetro):
            if DAPRetro >= 100000:
                return 'Needs to Confirm DAP Retro $ Amount'
            else:
                return ''
        data_xls ['Retro Award Tester'] = data_xls.apply(lambda x : RetroAwardTester(x['DAP Retroactive Award']), axis = 1)          
                
       
        
        #BlankOutcomeTester
                                
        def BlankOutcomeTester (DAPOutcome):
            if DAPOutcome == '':
               return 'Needs Outcome'
            else :
               return ''
        data_xls ['Blank Outcome Tester'] = data_xls.apply(lambda x : BlankOutcomeTester(x['DAP Outcome']), axis = 1)
        
        
        
        #If case has benefits, then it can't have short service as an outcome
        
        NoBenefitsList = ['Short or other services','Client did not receive / retain benefits','Client withdrew/failed to return','Client won/did not receive any benefits','Case Remanded']
        
        def DAPOutcomeTester (DAPOutcome,DAPRetro,DAPInterim,SSIMonthly,DIBMonthly,NoBenefitsList):
            if DAPOutcome in NoBenefitsList and DAPRetro > 0:
                return 'Needs Higher Level Outcome'
            elif DAPOutcome in NoBenefitsList and DAPInterim > 0:
                return 'Needs Higher Level Outcome'           
            elif DAPOutcome in NoBenefitsList and SSIMonthly > 0:
                return 'Needs Higher Level Outcome'
            elif DAPOutcome in NoBenefitsList and DIBMonthly > 0:
                return 'Needs Higher Level Outcome'
            else :
                return ''
        data_xls ['DAP Outcome Tester'] = data_xls.apply(lambda x : DAPOutcomeTester(x['DAP Outcome'],x['DAP Retroactive Award'],x['Interim Assistance'],x['Monthly SSI Award'],x['Monthly DIB Award'],NoBenefitsList), axis = 1)
        
        
        #if the outcome is monthly benefits then either DAP Monthly or SSI Monthly has to have $
        def MonthlyBenefitsTester  (DAPOutcome,DIBMonthly,SSIMonthly):
            DIBMonthly = int(DIBMonthly)
            SSIMonthly = int(SSIMonthly)
            
            if DAPOutcome == 'Client won/ received  retained monthly benefits' and DIBMonthly == 0 and SSIMonthly == 0:
                return 'Needs Monthly Benefits'
            else:
                return ''
        data_xls ['DAP Monthly Benefits Tester'] = data_xls.apply(lambda x : MonthlyBenefitsTester(x['DAP Outcome'],x['Monthly DIB Award'],x['Monthly SSI Award']),axis = 1)
       
       
        #If the outcome says you got retro, must enter retro award $
       
        def DAPRetroTester (DAPOutcome,DAPRetro):
            DAPRetro = int(DAPRetro)
            if DAPOutcome == 'Client won/received only retroactive benefits' and DAPRetro == 0:
                return "Needs Retro Award $"
        data_xls ['DAP Retro Tester'] = data_xls.apply(lambda x : DAPRetroTester(x['DAP Outcome'],x['DAP Retroactive Award']),axis = 1)
        
        
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
        data_xls ['No Benefits Tester'] = data_xls.apply(lambda x : NoBenefitsTester(x['DAP Outcome'],x['DAP Retroactive Award'],x['Interim Assistance'],x['Monthly SSI Award'],x['Monthly DIB Award'],NoBenefitsList), axis = 1)
        
        #Received SSI Tester
        
        def SSITester (ReceivedSSI,SSIMonthly):
            SSIMonthly = int(SSIMonthly)
            if ReceivedSSI == 'Yes' and SSIMonthly == 0:
                return 'Needs SSI Award Amount'
            elif ReceivedSSI == 'No' and SSIMonthly > 0:
                return 'Needs Received SSI? = Yes'
            else:
                return ''
                
        data_xls ['SSI Tester'] = data_xls.apply(lambda x : SSITester(x['Received SSI?'],x['Monthly SSI Award']),axis = 1)
                
        
        #Received DIB Tester
        
        def DIBTester (ReceivedDIB,DIBMonthly):
            DIBMonthly = int(DIBMonthly)
            if ReceivedDIB == 'Yes' and DIBMonthly == 0:
                return 'Needs DIB Award Amount'
            elif ReceivedDIB == 'No' and DIBMonthly > 0:
                return 'Needs Received DIB? = Yes'
            else:
                return ''
                
        data_xls ['DIB Tester'] = data_xls.apply(lambda x : DIBTester(x['Received DIB?'],x['Monthly DIB Award']),axis = 1)
        
        
        #Ordering Spreadsheet Correctly
        
        data_xls = data_xls[['Hyperlinked Case #','Assigned Branch/CC','Primary Advocate','Client Name',
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
        'Received DIB?','DIB Tester'
        
        ]]        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Sheet1',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
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
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "DAPCleanup " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>DAP Exception Report</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>DAP Exceptions Finder Tool</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Cleanup!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2288" target="_blank">"DAP Exceptions for Python Tool"</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to add case hyperlinks to.</li> 
    <li>Once you have identified this file, click ‘Cleanup!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>Note, the column with your case ID numbers in it must be titled "Matter/Case ID#" or "id" for this to work.</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
