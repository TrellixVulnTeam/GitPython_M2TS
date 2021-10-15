#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/AnnualComparer", methods=['GET', 'POST'])
def AnnualComparer():
    #upload file from computer
    if request.method == 'POST':
        print(request.files.getlist('file'))
        f = request.files.getlist('file')[0]
        g = request.files.getlist('file')[1]
        print(f)
        print(g)
        
        
        df1 = pd.read_excel(f,skiprows=2)
        df2 = pd.read_excel(g,skiprows=2)
        
        df = df1.append(df2, ignore_index=True)
        
        df['Year'] = df['Date Closed'].apply(lambda x: x[-4:])
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)    
        
        df['Unit'] = df.apply(lambda x : DataWizardTools.UnitSplitter(x['Legal Problem Code']),axis=1)
        
        
        Unit_pivot = pd.pivot_table(df,values=['Matter/Case ID#'], index=['Unit'],columns = ['Year'], aggfunc=lambda x: len(x.unique()))
        
        Close_pivot = pd.pivot_table(df,values=['Matter/Case ID#'], index=['Close Reason'],columns = ['Year'], aggfunc=lambda x: len(x.unique()))
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        Unit_pivot.to_excel(writer, sheet_name='Unit Pivot')
        Close_pivot.to_excel(writer, sheet_name='Close Pivot')
        
        df.to_excel(writer, sheet_name='Raw Data',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Raw Data']
        unitpivot = writer.sheets['Unit Pivot']
        closepivot = writer.sheets['Close Pivot']
        
        unitchart = workbook.add_chart({'type':'column'})
        closechart = workbook.add_chart({'type':'column'})
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        #assign new format to columns A
        worksheet.set_column('A:A',20,link_format)
        unitpivot.set_column('A:A',20)
        closepivot.set_column('A:A',40)
        
        UnitCategoriesRange="='Unit Pivot'!$A$4:$A$"+str(Unit_pivot.shape[0]+3)
        UnitSeriesOneRange="='Unit Pivot'!$B$4:$B$"+str(Unit_pivot.shape[0]+3)
        UnitSeriesTwoRange="='Unit Pivot'!$C$4:$C$"+str(Unit_pivot.shape[0]+3)
        
        CloseCategoriesRange="='Close Pivot'!$A$4:$A$"+str(Unit_pivot.shape[0]+3)
        CloseSeriesOneRange="='Close Pivot'!$B$4:$B$"+str(Unit_pivot.shape[0]+3)
        CloseSeriesTwoRange="='Close Pivot'!$C$4:$C$"+str(Unit_pivot.shape[0]+3)
        

        
        unitchart.set_y_axis({
            'name': 'Cases Closed',
            'name_font': {'size': 14, 'bold': True},
            })
        unitchart.set_x_axis({
            'name': 'Practice Area',
            'name_font': {'size': 14, 'bold': True},
            })
            
        unitchart.set_size({'width': 720, 'height': 576})
        
        unitchart.add_series({
            'categories': UnitCategoriesRange,
            'name': "2020",
            'values': UnitSeriesOneRange
        })
        unitchart.add_series({
            'name': "2021",
            'values': UnitSeriesTwoRange
        })
        
        
        
        closechart.set_y_axis({
            'name': 'Cases Closed',
            'name_font': {'size': 14, 'bold': True},
            })
        closechart.set_x_axis({
            'name': 'Close Reason',
            'name_font': {'size': 14, 'bold': True},
            })
            
        closechart.set_size({'width': 720, 'height': 576})
        
        closechart.add_series({
            'categories': CloseCategoriesRange,
            'name': "2020",
            'values': CloseSeriesOneRange
        })
        closechart.add_series({
            'name': "2021",
            'values': CloseSeriesTwoRange
        })
        
        
        unitpivot.set_row(0, None, None, {'hidden': True})
        unitpivot.insert_chart('D2', unitchart)
        closepivot.set_row(0, None, None, {'hidden': True})
        closepivot.insert_chart('D2', closechart)
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Comparison " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Annual Comparer</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Compare two spreadsheets:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file multiple =""><input type=submit value=Hyperlink!>
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
    
