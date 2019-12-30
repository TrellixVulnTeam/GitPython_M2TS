#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


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
            data_xls = pd.read_excel(f,skiprows=2)
        else:
            data_xls = pd.read_excel(f)
        

        #apply hyperlink methodology with splicing and concatenation
      
        def NoIDDelete(CaseID):
            if CaseID == '':
                return 'No Case ID'
            else:
                return str(CaseID)
        data_xls['Matter/Case ID#'] = data_xls.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = data_xls['Matter/Case ID#']
        data_xls['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del data_xls['Matter/Case ID#']
        move=data_xls['Temp Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        
        #split into separate dataframes
        
        BkLSdata_xls = data_xls[data_xls['Assigned Branch/CC'] == 'Brooklyn Legal Services']
        BxLSdata_xls = data_xls[data_xls['Assigned Branch/CC'] == 'Bronx Legal Services']
        QLSdata_xls = data_xls[data_xls['Assigned Branch/CC'] == 'Queens Legal Services']
        MLSdata_xls = data_xls[data_xls['Assigned Branch/CC'] == 'Manhattan Legal Services']
        SILSdata_xls = data_xls[data_xls['Assigned Branch/CC'] == 'Staten Island Legal Services']
        LSUdata_xls = data_xls[data_xls['Assigned Branch/CC'] == 'Legal Support Unit']
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        
        BkLSdata_xls.to_excel(writer, sheet_name='Brooklyn',index=False)
        QLSdata_xls.to_excel(writer, sheet_name='Queens', index=False)
        MLSdata_xls.to_excel(writer, sheet_name='Manhattan', index=False)
        BxLSdata_xls.to_excel(writer, sheet_name='Bronx', index=False)
        SILSdata_xls.to_excel(writer, sheet_name='Staten Island', index=False)
        LSUdata_xls.to_excel(writer, sheet_name='LSU', index=False)
        
        
        #rename worksheets with their original names 
        workbook = writer.book
        
        worksheetBkLS = writer.sheets['Brooklyn']
        worksheetQLS = writer.sheets['Queens']
        worksheetMLS = writer.sheets['Manhattan']
        worksheetBxLS = writer.sheets['Bronx']
        worksheetSILS = writer.sheets['Staten Island']
        worksheetLSU = writer.sheets['LSU']
        
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        
        worksheetBkLS.set_column('A:A',20,link_format)
        worksheetQLS.set_column('A:A',20,link_format)
        worksheetMLS.set_column('A:A',20,link_format)
        worksheetBxLS.set_column('A:A',20,link_format)
        worksheetSILS.set_column('A:A',20,link_format)
        worksheetLSU.set_column('A:A',20,link_format)
        worksheetBkLS.set_column('B:ZZ',25)
        worksheetQLS.set_column('B:ZZ',25)
        worksheetMLS.set_column('B:ZZ',25)
        worksheetBxLS.set_column('B:ZZ',25)
        worksheetSILS.set_column('B:ZZ',25)
        worksheetLSU.set_column('B:ZZ',25)
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Split " + f.filename)

        
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
    
