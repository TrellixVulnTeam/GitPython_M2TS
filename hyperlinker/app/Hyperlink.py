#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/upload", methods=['GET', 'POST'])
@app.route("/hyperlinker", methods=['GET', 'POST'])
def upload_hyperlink():
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
        
        #find excel worksheet names and make them into a dataframe
        sheet_names=pd.read_excel(f,sheet_name=None)
        worksheetnames  = pd.DataFrame(sheet_names.keys())
        number_of_sheets = len(worksheetnames.index)
        
        #load 9 additional sheets if they're there
        if number_of_sheets > 1:
            data_xls2 = pd.read_excel(f, sheet_name=1)
        
        if number_of_sheets > 2:
            data_xls3 = pd.read_excel(f, sheet_name=2)
            
        if number_of_sheets > 3:
            data_xls4 = pd.read_excel(f, sheet_name=3)
            
        if number_of_sheets > 4:
            data_xls5 = pd.read_excel(f, sheet_name=4)
        
        if number_of_sheets > 5:
            data_xls6 = pd.read_excel(f, sheet_name=5)
        
        if number_of_sheets > 6:
            data_xls7 = pd.read_excel(f, sheet_name=6)
        
        if number_of_sheets > 7:
            data_xls8 = pd.read_excel(f, sheet_name=7)
            
        if number_of_sheets > 8:
            data_xls9 = pd.read_excel(f, sheet_name=8)

        if number_of_sheets > 9:
            data_xls10 = pd.read_excel(f, sheet_name=9)        
        
        #if this is a report that has 'id' instead of 'Matter/Case id#' as the id column header - this will change it
        if 'Matter/Case ID#' not in data_xls.columns:
            data_xls['Matter/Case ID#'] = data_xls['id'].astype(str)
            del data_xls['id']
        data_xls['Matter/Case ID#'] = data_xls['Matter/Case ID#'].astype(str)
        
        #apply hyperlink methodology with splicing and concatenation
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[-7:])
        CaseNum = data_xls['Matter/Case ID#']
        data_xls['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del data_xls['Matter/Case ID#']
        move=data_xls['Temp Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        #FOR SECOND SHEET
                
        if number_of_sheets > 1:
            
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls2.columns:
                if 'id' in data_xls2.columns:
                    data_xls2['Matter/Case ID#'] = data_xls2['id']
                    del data_xls2['id']
                
            if 'Matter/Case ID#' in data_xls2.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls2['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls2['Matter/Case ID#']
                data_xls2['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls2['Matter/Case ID#']
                move=data_xls2['Temp Hyperlinked Case #']
                data_xls2.insert(0,'Hyperlinked Case #', move)           
                del data_xls2['Temp Hyperlinked Case #']
        
        
        #FOR Third SHEET
                
        if number_of_sheets > 2:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls3.columns:
                if 'id' in data_xls3.columns:
                    data_xls3['Matter/Case ID#'] = data_xls3['id']
                    del data_xls3['id']
                
            if 'Matter/Case ID#' in data_xls3.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls3['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls3['Matter/Case ID#']
                data_xls3['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls3['Matter/Case ID#']
                move=data_xls3['Temp Hyperlinked Case #']
                data_xls3.insert(0,'Hyperlinked Case #', move)           
                del data_xls3['Temp Hyperlinked Case #']
                
        #Fourth Sheet
        
        if number_of_sheets > 3:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls4.columns:
                if 'id' in data_xls4.columns:
                    data_xls4['Matter/Case ID#'] = data_xls4['id']
                    del data_xls4['id']
                
            if 'Matter/Case ID#' in data_xls4.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls4['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls4['Matter/Case ID#']
                data_xls4['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls4['Matter/Case ID#']
                move=data_xls4['Temp Hyperlinked Case #']
                data_xls4.insert(0,'Hyperlinked Case #', move)           
                del data_xls4['Temp Hyperlinked Case #']
        
        #Fifth Sheet
        
        if number_of_sheets > 4:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls5.columns:
                if 'id' in data_xls5.columns:
                    data_xls5['Matter/Case ID#'] = data_xls5['id']
                    del data_xls5['id']
                
            if 'Matter/Case ID#' in data_xls5.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls5['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls5['Matter/Case ID#']
                data_xls5['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls5['Matter/Case ID#']
                move=data_xls5['Temp Hyperlinked Case #']
                data_xls5.insert(0,'Hyperlinked Case #', move)           
                del data_xls5['Temp Hyperlinked Case #']
                
        #Sixth Sheet
        
        if number_of_sheets > 5:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls6.columns:
                if 'id' in data_xls6.columns:
                    data_xls6['Matter/Case ID#'] = data_xls6['id']
                    del data_xls6['id']
                
            if 'Matter/Case ID#' in data_xls6.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls6['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls6['Matter/Case ID#']
                data_xls6['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls6['Matter/Case ID#']
                move=data_xls6['Temp Hyperlinked Case #']
                data_xls6.insert(0,'Hyperlinked Case #', move)           
                del data_xls6['Temp Hyperlinked Case #']
                
        #Seventh Sheet
        
        if number_of_sheets > 6:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls7.columns:
                if 'id' in data_xls7.columns:
                    data_xls7['Matter/Case ID#'] = data_xls7['id']
                    del data_xls7['id']
                
            if 'Matter/Case ID#' in data_xls7.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls7['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls7['Matter/Case ID#']
                data_xls7['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls7['Matter/Case ID#']
                move=data_xls7['Temp Hyperlinked Case #']
                data_xls7.insert(0,'Hyperlinked Case #', move)           
                del data_xls7['Temp Hyperlinked Case #']
                
        #Eigth Sheet
        
        if number_of_sheets > 7:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls8.columns:
                if 'id' in data_xls8.columns:
                    data_xls8['Matter/Case ID#'] = data_xls8['id']
                    del data_xls8['id']
                
            if 'Matter/Case ID#' in data_xls8.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls8['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls8['Matter/Case ID#']
                data_xls8['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls8['Matter/Case ID#']
                move=data_xls8['Temp Hyperlinked Case #']
                data_xls8.insert(0,'Hyperlinked Case #', move)           
                del data_xls8['Temp Hyperlinked Case #']
                
        #Ninth Sheet
        
        if number_of_sheets > 8:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls9.columns:
                if 'id' in data_xls9.columns:
                    data_xls9['Matter/Case ID#'] = data_xls9['id']
                    del data_xls9['id']
                
            if 'Matter/Case ID#' in data_xls9.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls9['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls9['Matter/Case ID#']
                data_xls9['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls9['Matter/Case ID#']
                move=data_xls9['Temp Hyperlinked Case #']
                data_xls9.insert(0,'Hyperlinked Case #', move)           
                del data_xls9['Temp Hyperlinked Case #']
                
        #Tenth Sheet
        
        if number_of_sheets > 9:
        
            #if this is a report that has 'id' instead of 'matter/case id#' as the id column header - this will change it
            if 'Matter/Case ID#' not in data_xls10.columns:
                if 'id' in data_xls10.columns:
                    data_xls10['Matter/Case ID#'] = data_xls10['id']
                    del data_xls10['id']
                
            if 'Matter/Case ID#' in data_xls10.columns:    
        
                #apply hyperlink methodology with splicing and concatenation
                last7 = data_xls10['Matter/Case ID#'].apply(lambda x: x[-7:])
                CaseNum = data_xls10['Matter/Case ID#']
                data_xls10['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
                del data_xls10['Matter/Case ID#']
                move=data_xls10['Temp Hyperlinked Case #']
                data_xls10.insert(0,'Hyperlinked Case #', move)           
                del data_xls10['Temp Hyperlinked Case #']
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name=worksheetnames.iloc[0,0],index=False)
        if number_of_sheets > 1:
            data_xls2.to_excel(writer, sheet_name=worksheetnames.iloc[1,0],index=False)
        if number_of_sheets > 2:
            data_xls3.to_excel(writer, sheet_name=worksheetnames.iloc[2,0],index=False)
        if number_of_sheets > 3:
            data_xls4.to_excel(writer, sheet_name=worksheetnames.iloc[3,0],index=False)
        if number_of_sheets > 4:
            data_xls5.to_excel(writer, sheet_name=worksheetnames.iloc[4,0],index=False)
        if number_of_sheets > 5:
            data_xls6.to_excel(writer, sheet_name=worksheetnames.iloc[5,0],index=False)
        if number_of_sheets > 6:
            data_xls7.to_excel(writer, sheet_name=worksheetnames.iloc[6,0],index=False)
        if number_of_sheets > 7:
            data_xls8.to_excel(writer, sheet_name=worksheetnames.iloc[7,0],index=False)
        if number_of_sheets > 8:
            data_xls9.to_excel(writer, sheet_name=worksheetnames.iloc[8,0],index=False)
        if number_of_sheets > 9:
            data_xls10.to_excel(writer, sheet_name=worksheetnames.iloc[9,0],index=False)
        
        #rename worksheets with their original names 
        workbook = writer.book
        worksheet = writer.sheets[worksheetnames.iloc[0,0]]
        if number_of_sheets > 1:
            worksheet2 = writer.sheets[worksheetnames.iloc[1,0]]
        if number_of_sheets > 2:
            worksheet3 = writer.sheets[worksheetnames.iloc[2,0]]
        if number_of_sheets > 3:
            worksheet4 = writer.sheets[worksheetnames.iloc[3,0]]
        if number_of_sheets > 4:
            worksheet5 = writer.sheets[worksheetnames.iloc[4,0]]
        if number_of_sheets > 5:
            worksheet6 = writer.sheets[worksheetnames.iloc[5,0]]
        if number_of_sheets > 6:
            worksheet7 = writer.sheets[worksheetnames.iloc[6,0]]
        if number_of_sheets > 7:
            worksheet8 = writer.sheets[worksheetnames.iloc[7,0]]
        if number_of_sheets > 8:
            worksheet9 = writer.sheets[worksheetnames.iloc[8,0]]
        if number_of_sheets > 9:
            worksheet10 = writer.sheets[worksheetnames.iloc[9,0]]

        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})

        worksheet.set_column('A:A',20,link_format)
        if number_of_sheets > 1:
            worksheet2.set_column('A:A',20,link_format)
        if number_of_sheets > 2:
            worksheet3.set_column('A:A',20,link_format)
        if number_of_sheets > 3:
            worksheet4.set_column('A:A',20,link_format)
        if number_of_sheets > 4:
            worksheet5.set_column('A:A',20,link_format)
        if number_of_sheets > 5:
            worksheet6.set_column('A:A',20,link_format)
        if number_of_sheets > 6:
            worksheet7.set_column('A:A',20,link_format)
        if number_of_sheets > 7:
            worksheet8.set_column('A:A',20,link_format)
        if number_of_sheets > 8:
            worksheet9.set_column('A:A',20,link_format)
        if number_of_sheets > 9:
            worksheet10.set_column('A:A',20,link_format)

        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Hyperlinked " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Hyperlinks</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Add Hyperlinks to Your Report:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Hyperlink!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to add case hyperlinks to.</li> 
    <li>Once you have identified this file, click ‘Hyperlink!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>Note, the column with your case ID numbers in it must be titled "Matter/Case ID#" or "id" for this to work.</li>
    <li>This will work for excel documents with up to 10 separate 'worksheets' (additional worksheets must have the top two rows of a 'raw' LegalServer report removed).</li></ul>
    </br>
    <a href="/">Home</a>
    '''
    
