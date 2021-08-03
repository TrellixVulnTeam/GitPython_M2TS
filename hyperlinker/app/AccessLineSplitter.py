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
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']

        #Create Hyperlinks
        df['ClientID'] = df['Matter/Case ID#']
        df['Case ID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)  
        
        AccessLineStaff = [
        "Pierre, Haenley",
        "Ortega, Luis",
        "Djourab, Atteib",
        "Suriel, Sal",
        "Villanueva, Anthony",
        "Ruiz-Caceres, Gaby A",
        "Yeh, Victoria"]
        
        def SplittingFunction (CaseWorkerName):
            if CaseWorkerName in AccessLineStaff:
                return "Access Line"
            else:
                return "Not Access Line"
        
        df["Access Line?"] = df.apply(lambda x : SplittingFunction(x['Caseworker Name']),axis = 1)


        #Split into different tabs
        output_dictionary = dict(tuple(df.groupby('Access Line?')))
        
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
                ws.freeze_panes(1, 2)
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
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to split into different documents by borough.</li> 
    <li>Once you have identified this file, click ‘Split!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>In order for this tool to work your column header with intake paralegal in it needs to read as "Caseworker Name".</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
