from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/IOIemp", methods=['GET', 'POST'])
def upload_IOIemp():
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
        data_xls.insert(0,'Hyperlinked Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Bronx Legal Services','BxLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Brooklyn Legal Services','BkLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Queens Legal Services','QLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Manhattan Legal Services','MLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Staten Island Legal Services','SILS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Legal Support Unit','LSU')
        
        #Determining 'level of service' from 3 fields       
        def HRA_Service_Type(Employment_Tier):
            Employment_Tier = str(Employment_Tier)
            
            if Employment_Tier.startswith("Advice-No Retainer") == True:
                return 'B'
            elif Employment_Tier.startswith("UI Representation") == True:
                return 'T1'
            elif Employment_Tier.startswith("Advice-Investigation Retainer") == True:
                return 'T1'
            elif Employment_Tier.startswith("Demand Letter-Negotiation") == True:
                return 'T1'
            elif Employment_Tier.startswith("Admin Rep") == True:
                return 'T2'
            elif Employment_Tier.startswith("Litigation") == True:
                return 'T2'
            else:
                return '***Needs Cleanup***'
        data_xls['HRA Service Type'] = data_xls.apply(lambda x: HRA_Service_Type(x['HRA IOI Employment Law IOI Employment Tier Category:']), axis=1)

        data_xls['HRA Proceeding Type'] = 'EMP'
        
        def Income_Exclude(IncomePct,Waiver):
            IncomePct = int(IncomePct)
            Waiver = str(Waiver)
            if IncomePct > 200 and Waiver.startswith('Y') == False:
                return 'Needs Income Waiver'
            else:
                return ''
        

        data_xls['Exclude due to Income?'] = data_xls.apply(lambda x: Income_Exclude(x['Percentage of Poverty'],x['HRA IOI Employment Law If Client is Over 200% of FPL, Did you seek a waiver from HRA?']), axis=1)
        
        data_xls['Open Month'] = data_xls['Date Opened'].apply(lambda x: str(x)[:2])
        data_xls['Open Day'] = data_xls['Date Opened'].apply(lambda x: str(x)[3:5])
        data_xls['Open Year'] = data_xls['Date Opened'].apply(lambda x: str(x)[6:])
        data_xls['Open Construct'] = data_xls['Open Year'] + data_xls['Open Month'] + data_xls['Open Day']
        
        
        def DHCI_Needed(DHCI,Employment_Tier,Open_Construct):
            if Employment_Tier == 'Advice-No Retainer':
                return ''
            elif Open_Construct == 'na':
                return ''
            elif int(Open_Construct) < 20181115:
                return ''
            elif DHCI != 'Yes':
                return 'Needs DHCI Form'
            else:
                return ''
        
        data_xls['Needs DHCI?'] = data_xls.apply(lambda x: DHCI_Needed(x['HRA IOI Employment Law DHCI Form?'],x['HRA IOI Employment Law IOI Employment Tier Category:'],x['Open Construct']), axis=1)

        data_xls['Employment Tier Category'] = data_xls['HRA IOI Employment Law IOI Employment Tier Category:']
        
        data_xls['Client Name'] = data_xls['Full Person/Group Name (Last First)']
        
        data_xls['Office'] = data_xls['Assigned Branch/CC']
        
        #Put everything in the right order
        data_xls = data_xls[['Hyperlinked Case #','Office','Primary Advocate','Client Name','Employment Tier Category','Needs DHCI?','Exclude due to Income?','HRA Service Type']]
                      
        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Sheet1',index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})

        worksheet.set_column('A:A',20,link_format)

        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>IOI Employment</title>
    <h1>Check your IOI Employment Cases:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=IOI-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called "Grants Management IOI Employment (3474) Report".</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for IOI.</li> 
    <li>Once you have identified this file, click ‘IOI-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
