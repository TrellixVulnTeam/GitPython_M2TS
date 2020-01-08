#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/DHCIPrep", methods=['GET', 'POST'])
def DHCIPrep():
    #upload file from computer
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        #turn the excel file into a dataframe, but skip the top 2 rows if they are blank
        test = pd.read_excel(f)
        test.fillna('',inplace=True)
        if test.iloc[0][0] == '':
            data_xls = pd.read_excel(f,skiprows=2)
        else:
            data_xls = pd.read_excel(f)
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[-7:])
        data_xls['Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + data_xls['Matter/Case ID#'] +'"' +')'
        move = data_xls['Hyperlinked Case #']
        del data_xls['Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #',move)           
        
        
        data_xls = data_xls.fillna(' ')
        
        #Filter out families with children under 200.99 percent of poverty
        def ChildPovTester (Children, PovertyPercent):
            if Children > 0 and PovertyPercent <= 200.99:
                return 'Good'
            else:
                return 'No'
        data_xls['Child & Poverty Tester'] = data_xls.apply(lambda x: ChildPovTester(x['Number of People under 18'], x['Percentage of Poverty']), axis=1)
        
        
        #Test if social security number is correct format
        def SSNum (SSNum):
            SSNum = str(SSNum)
            First3 = SSNum[0:3]
            Middle2 = SSNum[4:6]
            Last4 = SSNum[7:11]
            FirstDash = SSNum[3:4]
            SecondDash = SSNum[6:7]
            
            if First3 == '000' and Middle2 == '00':
                return 'No'
            elif str.isnumeric(First3) == True and str.isnumeric(Middle2) == True and str.isnumeric(Last4) == True and FirstDash == '-' and SecondDash == '-': 
                return 'Good'
            else:
                return "No"
                
        data_xls['SS Tester'] = data_xls.apply(lambda x: SSNum(x['Social Security #']), axis=1)
        
        #Test that PA isn't there
        def PATester (PANum):
            PANum = str(PANum).upper()
            if PANum == ' ' or PANum.startswith('NO') or PANum.startswith('UNK') or PANum.startswith('UNA') or PANum.startswith('WILL') or PANum.startswith('N/A') or PANum.startswith('000') or PANum.startswith('CLIENT') or PANum.startswith('NEED'):
                return 'Good'
            else:
                return 'No'
        data_xls['PA Tester'] = data_xls.apply(lambda x: PATester(x["Gen Pub Assist Case Number"]), axis=1)
        
        def DHCIFilter (PATester,SSTester,ChildPovTester):
            if PATester == 'Good' and SSTester == 'Good' and ChildPovTester == 'Good':
                return 'Ready for Extractor!'
            else:
                return ''
                
        data_xls['Extractor Ready?'] = data_xls.apply(lambda x: DHCIFilter(x["PA Tester"],x["SS Tester"],x["Child & Poverty Tester"]), axis=1)
        
        data_xls['Client Full Name'] = data_xls['Client First Name'] + ' ' + data_xls['Client Last Name']
        
        ForPrep_xls = data_xls[data_xls['Extractor Ready?'] == 'Ready for Extractor!']
        
        data_xls = data_xls[['Hyperlinked Case #',
        'Primary Advocate',
        "Assigned Branch/CC",
        'Child & Poverty Tester',
        "Number of People 18 and Over",
        "Number of People under 18",
        "Percentage of Poverty",
        "Total Annual Income ",
        "SS Tester",
        "Social Security #",
        "PA Tester",
        "Gen Pub Assist Case Number",
        'Extractor Ready?',
        "Client Full Name",
        "Street Address",
        "Apt#/Suite#",
        "City",
        "Zip Code"]]
        
        ForPrep_xls = ForPrep_xls[['Matter/Case ID#', 
        "Client Full Name",
        "Assigned Branch/CC"]]
        
        
        
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Sheet1',index=False)
        ForPrep_xls.to_excel(writer, sheet_name = 'Prepped for Extraction',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet2 = writer.sheets['Prepped for Extraction']
        
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        worksheet.set_column('A:A',20,link_format)
        worksheet.set_column('B:ZZ',30)
        worksheet2.set_column('A:C',30)
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "DHCI Prepped " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>DHCI Extractor Prep</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Prepare Relevant Cases for DHCI Extraction:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Prep!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1507" target="_blank">TRC Raw Case Data Report</a>.</li> 
    <li>It will identify cases that have children in the household, are under 200.99% of the poverty level, do not have a PA#, and have full social security numbers. These cases will be placed in a second 'sheet' in the excel document that is formatted for the DHCI document extraction tool.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
