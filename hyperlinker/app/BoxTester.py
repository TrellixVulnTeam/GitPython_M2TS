#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os



@app.route("/BoxTester", methods=['GET', 'POST'])
def BoxTester():
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
      
        def NoIDDelete(CaseID):
            if CaseID == '':
                return 'No Case ID'
            else:
                return str(CaseID)
        df['Matter/Case ID#'] = df.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        
        last7 = df['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = df['Matter/Case ID#']
        df['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del df['Matter/Case ID#']
        move=df['Temp Hyperlinked Case #']
        df.insert(0,'Hyperlinked Case #', move)           
        del df['Temp Hyperlinked Case #']
        
        """
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)    
        """
        print(request.form['name'])
        print(request.form['date'])
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        highlight_link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True, 'bg_color':'yellow'})
        
        #assign new format to column A
        worksheet.set_column('A:A',20,link_format)
        
        if request.form.get('highlights'):
            worksheet.set_column('A:A',20,highlight_link_format)
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Hyperlinked " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Hyperlinks</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Checkbox Changes Script Behavior:</h1>
    <form action="" method=post enctype=multipart/form-data>
    
    <p><input type=file name=file><input type=submit value=Hyperlink!>
    
    </br>
    </br>
    <input type="checkbox" id="highlights" name="highlights" value="highlights" checked>
    <label for="highlights"> Add highlights</label><br><br>
    
    <input type = "text" id="name" name="name">
    <label for ="name"> Add some text!</label><br><br>
    
    <input type = "date" id="date" name="date">
    <label for = "date"> choose a date! </label>
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to add case hyperlinks to.</li> 
    <li>Once you have identified this file, click ‘Hyperlink!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>Note, the column with your case ID numbers in it must be titled "Matter/Case ID#" or "id" for this to work.</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
