#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os
from zipfile import ZipFile


@app.route("/BoroughSplitter", methods=['GET', 'POST'])
def split_by_borough():
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
        

        #delete blank rows
      
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #hyperlink
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)    
        
        #Preparing Excel Document

        
        
        #define function, with single variable of dictionary created immediately above
        def save_xls(ExcelSplit, TabSplit):
            #Create dictionary of dataframes, each named after each unique value of 'assigned branch/cc'
            excel_dict_df = dict(tuple(df.groupby(ExcelSplit)))
            #use 'ZipFile' create empty zip folder, and assign 'newzip' as function-calling name
            with ZipFile("app\\sheets\\zipped\\SplitBoroughs.zip","w") as newzip:
                
                #starts cycling through each dataframe (each borough's data)
                for i in excel_dict_df:
                    
                    #because this is in for loop it creates a new excel file for each 'i' (i.e. each borough)
                    writer = pd.ExcelWriter(path = "app\\sheets\\zipped\\" + i + ".xlsx", engine = 'xlsxwriter')
                    
                    #create a dictionary of dataframes for each unique advocate, within the borough that our 'i' for loop is cycling through
                    tab_dict_df = dict(tuple(excel_dict_df[i].groupby(TabSplit)))
                    
                    #start cycling through each of these new advocate-based dataframes
                    for j in tab_dict_df:
                        
                        #write a tab in the borough's excel sheet, composed just of the advocate's cases, with the advocate's name
                        tab_dict_df[j].to_excel(writer, j, index = False)
                        
                        #creates ability to format tabs
                        worksheet = writer.sheets[j]
                        
                        #creates and adds formatting to tab
                        workbook = writer.book
                        link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                        worksheet.set_column('A:A',15,link_format)
                        worksheet.set_column('B:ZZ',20)
                        
                    #save's excel file
                    writer.save()
                    #adds excel file to zipped folder
                    newzip.write("app\\sheets\\zipped\\" + i + ".xlsx",arcname = i + ".xlsx")
          
        save_xls(ExcelSplit = 'Assigned Branch/CC', TabSplit = 'Caseworker Name')
        
        
            
            #newzip.write("app\\sheets\\zipped\\Brooklyn Legal Services ZipSplitterTester.xlsx", arcname = "Brooklyn Legal ServicesZipSplitterTester.xlsx")
            #newzip.write("app\\sheets\\zipped\\Queens Legal Services ZipSplitterTester.xlsx", arcname = "Queens Legal ServicesZipSplitterTester.xlsx")
            
        
        return send_from_directory('sheets\\zipped','SplitBoroughs.zip', as_attachment = True)
        
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Borough Splitter</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>Split Your Spreadsheet by Office:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Split!>
    </form>
    
    
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to split into different documents by borough.</li> 
    <li>Once you have identified this file, click ‘Split!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>In order for this tool to work your column header with boroughs in it needs to read as "Assigned Branch/CC" and the borough names must be in standard LegalServer format.</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
