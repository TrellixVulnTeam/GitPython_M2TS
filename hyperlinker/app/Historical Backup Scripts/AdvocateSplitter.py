#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/AdvocateSplitter", methods=['GET', 'POST'])
def AdvocateSplitter():
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

        if 'Matter/Case ID#' not in data_xls.columns:
            data_xls['Matter/Case ID#'] = data_xls['id']
            
        
        #separate last 7 digits from Case ID# so that it can be used for the link
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[-7:])

        #create excel formula with hyperlinks
        data_xls['Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + data_xls['Matter/Case ID#'] +'"' +')'
        
        #this is a way to move a newly created column to the very front of the spreadsheet
        move = data_xls['Hyperlinked Case #']
        del data_xls['Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #',move)           
        
        #delete original case number column since it's no longer necessary
        del data_xls['Matter/Case ID#']
        
        
        advocate_dictionary = dict(tuple(data_xls.groupby('Primary Advocate')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                ws = writer.sheets[i]
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = advocate_dictionary, path = "app\\sheets\\" + output_filename)
        
        
        """
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Sheet1',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        #assign new format to column A
        worksheet.set_column('A:A',20,link_format)
        
        writer.save()
        
        """
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Split " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Advocate Splitter</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Split Report by Advocates:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Split!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to add case hyperlinks to.</li> 
    <li>Your spreadsheet must have a columns labelled "Primary Advocate", and "Matter/Case ID#"  for this tool to work.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
