from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/TRCclean", methods=['GET', 'POST'])
def upload_TRCclean():
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
        
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = data_xls['Matter/Case ID#']
        data_xls['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del data_xls['Matter/Case ID#']
        move=data_xls['Temp Hyperlinked Case #']
        data_xls.insert(0,'Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Bronx Legal Services','BxLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Brooklyn Legal Services','BkLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Queens Legal Services','QLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Manhattan Legal Services','MLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Staten Island Legal Services','SILS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Legal Support Unit','LSU')
        
        #Needs HRA Release?
        
        def Release_Needed(HRA_Release):
            if HRA_Release == 'Yes':
                return ''
            else:
                return 'Needs HRA Release'
        
        data_xls['Needs Release?'] = data_xls.apply(lambda x: Release_Needed(x['HRA Release?']), axis=1)
        
        #Needs Level of Service?
        
        def Needs_LOS(Level_of_Service):
            if Level_of_Service == 'Advice' or Level_of_Service == 'Hold For Review' or Level_of_Service == 'Brief Service' or Level_of_Service == 'Representation - State Court' or Level_of_Service == 'Out-of-Court Advocacy' or Level_of_Service == 'Representation - Admin. Agency' or Level_of_Service == 'Representation - Federal Court'  :
                return ''
            else:
                return 'Needs Level of Service'
        
        data_xls['Needs Level of Service?'] = data_xls.apply(lambda x: Needs_LOS(x['Housing Level of Service']), axis=1)
        
        #Eligibility Date Within Current Fiscal Year?
        
        data_xls['Eligibility Month'] = data_xls['HAL Eligibility Date'].apply(lambda x: str(x)[:2])
        data_xls['Eligibility Day'] = data_xls['HAL Eligibility Date'].apply(lambda x: str(x)[3:5])
        data_xls['Eligibility Year'] = data_xls['HAL Eligibility Date'].apply(lambda x: str(x)[6:])
        data_xls['Eligibility Construct'] = data_xls['Eligibility Year'] + data_xls['Eligibility Month'] + data_xls['Eligibility Day']
        
        
        def Current_Year(Eligibility_Construct):
            if Eligibility_Construct == 'na':
                return 'No Eligibility Date'            
            elif int(Eligibility_Construct) > 20180701 and int(Eligibility_Construct) < 20190630:
                return ''
            else:
                return 'Wrong Fiscal Year'
        
        data_xls['Eligibility Date in Current Fiscal Year?'] = data_xls.apply(lambda x: Current_Year(x['Eligibility Construct']), axis=1)
        
        #Income Verified
        
        def Income_Verified(DHCI_form,Housing_Verification):
            if DHCI_form == 'Yes':
                return ''
            elif Housing_Verification == 'DHCI Form' or Housing_Verification == 'Active CA/SNAP':
                return ''
            else:
                return 'Needs Income Verified'
        
        data_xls['Income Verified?'] = data_xls.apply(lambda x: Income_Verified(x['Housing Signed DHCI Form'], x['Housing Income Verification']), axis=1)
        
        #Social Security or Public Assistance #
        
        def ValidSSN(SSN):
            if len(str(SSN)) < 11:
                return 'Invalid SSN'
            elif SSN == '000-00-0000' or SSN == '999-99-9999':
                return 'Invalid SSN'
            else:
                return ''
            
        data_xls['Valid SSN?'] = data_xls.apply(lambda x: ValidSSN(x['Social Security #']),axis=1)
        
        def ValidPA(PA):
            if len(str(PA)) < 8:
                return 'Invalid PA#'
            elif PA == '000-00-0000':
                return 'Invalid PA#'
            else:
                return ''
            
        data_xls['Valid PA#?'] = data_xls.apply(lambda x: ValidPA(x['Gen Pub Assist Case Number']),axis=1)
        
        def NeedsSSNorPA(ValidSSN,ValidPA):
            if ValidSSN == 'Invalid SSN' and ValidPA == 'Invalid PA#':
                return 'Needs SSN or PA#'
            else:
                return ''
                
        data_xls['Needs SSN or PA#?'] = data_xls.apply(lambda x: NeedsSSNorPA(x['Valid SSN?'],x['Valid PA#?']),axis=1)
        
        #Do cases have any problems?
        
        def AnyProblems(NeedsRelease,EligibilityDate,IncomeVerified,NeedsSSNorPA,NeedsLoS):
            if NeedsRelease == '' and EligibilityDate == '' and IncomeVerified == '' and NeedsSSNorPA == '' and NeedsLoS == '':
                return 'No Problems'
            else:
                return 'Case Has Problems'
        
        data_xls['Any Problems?'] = data_xls.apply(lambda x: AnyProblems(x['Needs Release?'],x['Eligibility Date in Current Fiscal Year?'],x['Income Verified?'],x['Needs SSN or PA#?'],x['Needs Level of Service?']),axis=1)
        
        #replace commas with semicolons - currently unused
        data_xls['Housing Activity Indicators'] = data_xls['Housing Activity Indicators'].str.replace(',',';')
        
        data_xls['Housing Services Rendered to Client'] = data_xls['Housing Services Rendered to Client'].str.replace(',',';')
        
        #Put everything in the right order
        data_xls = data_xls[['Case #','Assigned Branch/CC','Needs Release?','Eligibility Date in Current Fiscal Year?','Income Verified?','Needs SSN or PA#?','Needs Level of Service?','Any Problems?']]
        
        #Construct Summary Tables
        problems_pivot = pd.pivot_table(data_xls,index=['Assigned Branch/CC'],values=['Case #'], columns=['Any Problems?'],aggfunc=len,fill_value=0)
        
        release_pivot = pd.pivot_table(data_xls,index=['Assigned Branch/CC'],values=['Case #'], columns=['Needs Release?'],aggfunc=len,fill_value=0)
        
        LOS_pivot = pd.pivot_table(data_xls,index=['Assigned Branch/CC'],values=['Case #'], columns=['Needs Level of Service?'],aggfunc=len,fill_value=0)
        
        income_pivot = pd.pivot_table(data_xls,index=['Assigned Branch/CC'],values=['Case #'],columns=['Income Verified?'],aggfunc=len,fill_value=0)
        
        eligibility_pivot = pd.pivot_table(data_xls,index=['Assigned Branch/CC'],values=['Case #'],columns=['Eligibility Date in Current Fiscal Year?'],aggfunc=len,fill_value=0)
        
        SSNPA_pivot = pd.pivot_table(data_xls,index=['Assigned Branch/CC'],values=['Case #'],columns=['Needs SSN or PA#?'],aggfunc=len,fill_value=0)
        
        #Bounce to a new excel file        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Case List',index=False)
        problems_pivot.to_excel(writer, sheet_name='Cases Summary')
        release_pivot.to_excel(writer, sheet_name='Needs Release Summary')
        income_pivot.to_excel(writer, sheet_name='Income Verified Summary')
        eligibility_pivot.to_excel(writer, sheet_name='Eligibility Date Summary')
        SSNPA_pivot.to_excel(writer, sheet_name='SSA PA Summary')
        LOS_pivot.to_excel(writer, sheet_name='Level of Service Summary')

        workbook = writer.book
        CaseList = writer.sheets['Case List']
        CaseSummary = writer.sheets['Cases Summary']
        ReleaseSummary = writer.sheets['Needs Release Summary']
        IncomeSummary = writer.sheets['Income Verified Summary']
        EligibilitySummary = writer.sheets['Eligibility Date Summary']
        SSAPASummary  = writer.sheets['SSA PA Summary']
        LOSSummary = writer.sheets['Level of Service Summary']

        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})

        CaseList.set_column('A:A',20,link_format)
        CaseList.set_column('B:H',22)
        CaseSummary.set_column('A:C',20)
        ReleaseSummary.set_column('A:C',20)
        IncomeSummary.set_column('A:C',20)
        CaseSummary.set_column('A:C',20)
        EligibilitySummary.set_column('A:A',34)
        EligibilitySummary.set_column('B:D',20)
        SSAPASummary.set_column('A:C',20)
        LOSSummary.set_column('A:C',20)

        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>TRC Cleaner</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Check for Problems in TRC cases:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=TRC-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1507" target="_blank">TRC Raw Case Data Report</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for TRC.</li> 
    <li>Once you have identified this file, click ‘TRC-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
