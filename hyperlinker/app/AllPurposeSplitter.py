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


@app.route("/AllPurposeSplitter", methods=['GET', 'POST'])
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

        #Add Agency column
        df['Agency'] = "LSNYC"
        
        print(request.form['SplitSpreadsheetsBy2'])
        
        #Preparing Excel Document
        if request.form['SplitSpreadsheetsBy'] == "Goosebumps":
            ExcelSplitChoice = request.form['SplitSpreadsheetsBy2']
        else:
            ExcelSplitChoice = request.form['SplitSpreadsheetsBy']
            
        if request.form['SplitCategory'] == "Goosebumps":
            TabSplitChoice = request.form['SplitCategory2']
        else:
            TabSplitChoice = request.form['SplitCategory']

        if request.form['SplitSpreadsheetsBy'] != "Agency":
            newzip = ZipFile("app\\sheets\\zipped\\Zipped " + f.filename[:-5]+".zip","w")
        
        #define function, with single variable of dictionary created immediately above
        def save_xls(ExcelSplit, TabSplit):
            #Create dictionary of dataframes, each named after each unique value of 'assigned branch/cc'
            excel_dict_df = dict(tuple(df.groupby(ExcelSplit)))
            #use 'ZipFile' create empty zip folder, and assign 'newzip' as function-calling name
            #with ZipFile("app\\sheets\\zipped\\SplitBoroughs.zip","w") as newzip:
                
            #starts cycling through each dataframe (each borough's data)
            for i in excel_dict_df:
                
                #because this is in for loop it creates a new excel file for each 'i' (i.e. each borough)
                global output_filename
                output_filename = i + " " + f.filename
                writer = pd.ExcelWriter("app\\sheets\\" + output_filename, engine = 'xlsxwriter')
                #writer = pd.ExcelWriter(path = "app\\sheets\\zipped\\" + i + ".xlsx", engine = 'xlsxwriter')
                
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
                if request.form['SplitSpreadsheetsBy'] != "Agency":
                #adds excel file to zipped folder
                    newzip.write("app\\sheets\\" + output_filename,arcname = output_filename)
                    #adds excel file to zipped folder
                    #newzip.write("app\\sheets\\zipped\\" + i + ".xlsx",arcname = i + ".xlsx")
              
        if request.form['SplitSpreadsheetsBy'] != "Agency":
            save_xls(ExcelSplit = ExcelSplitChoice, TabSplit = TabSplitChoice)
            newzip.close()
            return send_from_directory('sheets\\zipped', "Zipped " + f.filename[:-5]+".zip", as_attachment = True)
        else:
            #output_filename = f.filename
            save_xls(ExcelSplit = 'Agency', TabSplit = TabSplitChoice)
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = output_filename)
        #save_xls(ExcelSplit = 'Assigned Branch/CC', TabSplit = 'Caseworker Name')
        
        
            
            #newzip.write("app\\sheets\\zipped\\Brooklyn Legal Services ZipSplitterTester.xlsx", arcname = "Brooklyn Legal ServicesZipSplitterTester.xlsx")
            #newzip.write("app\\sheets\\zipped\\Queens Legal Services ZipSplitterTester.xlsx", arcname = "Queens Legal ServicesZipSplitterTester.xlsx")
            
        
        return send_from_directory('sheets\\zipped','SplitBoroughs.zip', as_attachment = True)
        
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>All Purpose Splitter</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>Split Your Spreadsheet as You Like:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Split!>
    
     </br>  
     </br> 
     <li>Excel Split</li> 
     <label for="SplitSpreadsheetsBy">Choose a split category from the dropdown list:</label>
     <select id="SplitSpreadsheetsBy" name="SplitSpreadsheetsBy">
      <option value="Goosebumps">Or write one in below</option>
      <option value="Agency">No Split</option>
      <option value="Assigned Branch/CC">Assigned Branch/CC</option>
      <option value="Primary Advocate">Primary Advocate</option>
      <option value="Primary Funding Codes">Primary Funding Codes</option>
      <option value="Legal Problem Code Category">Legal Problem Code Category</option>
            
     </select>
     
     </br>
     <input type = "text" id="SplitSpreadsheetsBy2" name="SplitSpreadsheetsBy2">
    <label for ="SplitSpreadsheetsBy2"> Type column name here</label><br><br>
    
    </br>
      
    <li>Tab Split</li> 
    <label for="SplitCategory">Choose a split category from the dropdown list:</label>
    <select id="SplitCategory" name="SplitCategory">
      <option value="Goosebumps">Or write one in below</option>
      <option value="Agency">No Split</option>
      <option value="Assigned Branch/CC">Assigned Branch/CC</option>
      <option value="Primary Advocate">Primary Advocate</option>
      <option value="Primary Funding Codes">Primary Funding Codes</option>
      <option value="Legal Problem Code Category">Legal Problem Code Category</option>
     </select> 
     </br> 
     
     <input type = "text" id="SplitCategory2" name="SplitCategory2">
    <label for ="SplitCategory2"> Type column name here</label><br><br>
    
    </br>
    
    </form>
    
    
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>Browse your computer using the "Browse" button above to find the LegalServer excel document that you want to split into different workbooks and/or tabs.</li> 
    <li>Once you have chosen this file, use the dropdown menu under "Excel Split" to choose the category you want to use to split your workbook. Use the "No Split" option if you don't want to split into multiple workbooks.</li>
    <li>If the option you want is not listed, type the column name exactly as it appears on your spreadsheet into the text box.</li>
    <li>If you'd like to split the tabs on your workbook, follow the same steps in the Tab Split section</li>    
    <li>Click ‘Split!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>In order for this tool to work your column header with boroughs in it needs to read as "Assigned Branch/CC" and the borough names must be in standard LegalServer format.</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
