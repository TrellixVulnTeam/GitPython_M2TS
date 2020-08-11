from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/TRCDuplicateIdentifier", methods=['GET', 'POST'])
def upload_TRCDuplicateIdentifier():
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
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)          
        
        df['Assigned Branch/CC'] = df.apply(lambda x : DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis=1)   


        
        
        #DuplicateTester
        #add client name and birth year and eligiblity date into one ID string
        #identify duplicates based on ID string
        #make new column identifying repeat values
        
        df['DupEligID'] = df["Client First Name"]+df["Client Last Name"]+df["Date of Birth"] +df["HAL Eligibility Date"]
        df['DuplicatedClient&EligDate?Bool'] = df.duplicated(['DupEligID'])
        
        def DuplicateHasEligDate (DupBool,EligDate):
            if EligDate == '':
                return False
            else:
                return DupBool
        df['DuplicatedClient&EligDate?Bool'] = df.apply(lambda x: DuplicateHasEligDate(x['DuplicatedClient&EligDate?Bool'], x['HAL Eligibility Date']), axis=1)
        
       

        dfs = df.groupby('DupEligID',sort = False)

        tdf = pd.DataFrame()
        for x, y in dfs:
            for z in y['DuplicatedClient&EligDate?Bool']:
                if z == True:
                    y['Duplicate Tester'] = 'Duplicate Found'
                else:
                    y['Duplicate Tester'] = 'No Duplicate Found'
                    
            tdf = tdf.append(y)
        df = tdf
        
       
       #***make it so that duplicates only show up if there's an eligibility date
    
        #Is everything okay with a case? 

        
            
       
        
        
        #sort by DupID
        
        df = df.sort_values(by=['DupEligID'])
        
        
        
        #Put everything in the right order
        
        df = df[[
        'Hyperlinked CaseID#',
        'Primary Advocate',
        'Duplicate Tester',
        "Date Opened",
        "Date Closed",
        "Client First Name",
        "Client Last Name",
       
        "HAL Eligibility Date",
        
        "Date of Birth",
        
        'DupEligID',
        
        
        "Assigned Branch/CC",
        
        
        
        ]]      
        
        #Preparing Excel Document
        
        #Split into different tabs
        allgood_dictionary = dict(tuple(df.groupby('Duplicate Tester')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                ws = writer.sheets[i]
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                regular_format = workbook.add_format({'font_color':'black'})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                bad_problem_format = workbook.add_format({'bg_color':'red'})
                medium_problem_format = workbook.add_format({'bg_color':'orange'})
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
                ws.autofilter('B1:ZZ1')
                ws.freeze_panes(1, 2)
                ws.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'No Release - Remove Elig Date',
                                                 'format': bad_problem_format})
                ws.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
                ws.conditional_format('C1:BO1',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Tester',
                                                 'format': problem_format})
                ws.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Must Have DHCI or PA#',
                                                 'format': medium_problem_format})            
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename)
       
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>TRC Duplicate Identifier</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>TRC Duplicate Identifier:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Duplicates?>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1507" target="_blank">TRC Raw Case Data Report</a>.</li>
    
    
    </br>
    <a href="/">Home</a>
    '''
