#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/DuplicateFinder", methods=['GET', 'POST'])
def DuplicateFinder():
    #upload file from computer
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        #turn the excel file into a dataframe, but skip the top 2 rows if they are blank
        test = pd.read_excel(f)
        test.fillna('',inplace=True)
        if test.iloc[0][0] == '':
            df = pd.read_excel(f,skiprows=2)
            print("Skipped top two rows")
        else:
            df = pd.read_excel(f)
            print("Dataframe starts from top")
        
        
        #apply hyperlink methodology with splicing and concatenation
      
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)    
        
        #create unique identifier:
        
        df['DupEligID'] = df["First Name"]+df["Last Name"]+df["Legal Problem Code"] +df["Client Date of Birth"] + df['Close Reason'] + str(df['Client Social Security #'])
        
        df['DuplicatedClient&EligDate?Bool'] = df.duplicated(['DupEligID'])
        

        dfs = df.groupby('DupEligID',sort = False)

        tdf = pd.DataFrame()
        for x, y in dfs:
            for z in y['DuplicatedClient&EligDate?Bool']:
                if z == True:
                    y['Duplicate Tester'] = 'Duplicate Found'
                else:
                    y['Duplicate Tester'] = ''
                    
            tdf = tdf.append(y)
        df = tdf
        
        df = df.sort_values(by=['DupEligID'])
        
        df = df[[
        'Hyperlinked CaseID#',
        'Duplicate Tester',
        'First Name',
        'Last Name',
        'Client Date of Birth',
        'Legal Problem Code',
        'Special Legal Problem Code',
        'Close Reason']]
        
        
        
        
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        #assign new format to column A
        
        
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(),len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx,col_idx, column_width)
        
        worksheet.set_column('A:A',20,link_format)
        
        
        problem_format = workbook.add_format({'bg_color':'yellow'})
        worksheet.conditional_format('B2:B100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Duplicate',
                                                 'format': problem_format})
        
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Hyperlinked " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Duplicate Finder</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Find Cases that LSC Would Consider Duplicates:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Hyperlink!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2455" target="_blank">Duplicate Case Finder</a>.</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
