#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/AccessLineSplitter", methods=['GET', 'POST'])
def AccessLineSplitter():
    #upload file from computer
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        #Test if there are extra rows at the top from a raw legalserver report - skip them if so     
        test = pd.read_excel(f)
        test.fillna('',inplace=True)
        if test.iloc[0][0] == '':
            df = pd.read_excel(f,skiprows=2)
        else:
            df = pd.read_excel(f)
        
        
        #Import Access Line Spreadsheet and turn it into dictionary with caseworkername:[start date, end date ]
        ALdf = pd.read_excel("app\\referencesheets\\Hotline Para List.xlsx")
        ALdf['End Date'] = ALdf['End Date'].fillna('2099-99-99')
        ALdf['Start Date 2'] = ALdf['Start Date 2'].fillna('2099-99-99')
        ALdf['End Date 2'] = ALdf['End Date 2'].fillna('2099-99-99')
        
        ALdf['Start Date'] = ALdf.apply(lambda x : DataWizardTools.DateMaker(x['Start Date']),axis=1) 
        ALdf['End Date'] = ALdf.apply(lambda x : DataWizardTools.DateMaker(x['End Date']),axis=1)
        ALdf['Start Date 2'] = ALdf.apply(lambda x : DataWizardTools.DateMaker(x['Start Date 2']),axis=1) 
        ALdf['End Date 2'] = ALdf.apply(lambda x : DataWizardTools.DateMaker(x['End Date 2']),axis=1)         
        
        
        AccessLineDictionary = ALdf.set_index('Name').T.to_dict('list')
        print(AccessLineDictionary)
        
        
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        
        
        #Create Hyperlinks
       
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)  
        
        
        if "Intake User" in df:
            df['Caseworker Name'] = df['Intake User']
        
        
        #Split cases in access line:
        df['Intake Date Construct'] = df.apply(lambda x : DataWizardTools.DateMaker(x['Intake Date']),axis=1) 
        
        def AccessLineSplit(CaseWorker,DateOpened):
            if CaseWorker in AccessLineDictionary:
                if AccessLineDictionary[CaseWorker][1] >= DateOpened >= AccessLineDictionary[CaseWorker][0]:
                    return 'Access Line'
                elif AccessLineDictionary[CaseWorker][3] >= DateOpened >= AccessLineDictionary[CaseWorker][2]:
                    return 'Access Line'
                else:
                    return 'Borough'
            else:
                return 'Borough'


        df['Access Line Case?'] = df.apply(lambda x: AccessLineSplit(x['Caseworker Name'],x['Intake Date Construct']),axis=1)     
        
        df = df.drop(['Intake Date Construct'], axis=1)
        
        #Split into different tabs
        output_dictionary = dict(tuple(df.groupby('Access Line Case?')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                ws = writer.sheets[i]
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                regular_format = workbook.add_format({'font_color':'black'})
                
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
                ws.autofilter('B1:ZZ1')
                ws.freeze_panes(1, 1)
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = output_dictionary, path = "app\\sheets\\" + output_filename)
       
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Split " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Access Line Splitter</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>Split Your Spreadsheet by whether it was opened by Access Line Staff:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Split!>
    </form>
    
    
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>This tool was designed to be used with the LegalServer Report <a href="https://lsnyc.legalserver.org/report/dynamic?load=2422" target="_blank">"Access Line Splitter Tool Report"</a> 
    </br>
    <a href="/">Home</a>
    '''
    
