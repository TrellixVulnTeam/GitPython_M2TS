#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, ImmigrationToolBox, EmploymentToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
from datetime import date
import pandas as pd
import os


@app.route("/IOIWaiverMaker", methods=['GET', 'POST'])
def IOIWaiverMaker():
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
        
        
        
        #Add Provider column #1
        df['Provider'] = 'LSNYC'
        
        #Add Contact Person column #2
        df['Contact Person'] = ''
        
        #Add Date of Request column #3
        df['Date of Request'] = date.today().strftime("%m/%d/%Y")
        
        #Add First & Last Initials column #4
        df['First & Last Initials'] = df['Client First Name'].apply(lambda x: x[0])+df['Client Last Name'].apply(lambda x: x[0])
                
        #Add New/Reconsideration? column #5
        df['New/Reconsideration?'] = 'New'
        
        #Add Income/Tier Upgrade/Other column #6
        df['Income/Tier Upgrade/Other'] = 'Income'
                            
        #Add Proceeding Type column #7 -part 1 special legal prob code?
        if request.form.get('Immigration'):
            print('Immigration Version')
            #Determining 'level of service' from 3 fields       
            df['HRA Level of Service'] = df.apply(lambda x: ImmigrationToolBox.HRA_Level_Service(x['Close Reason'],x['Level of Service']), axis=1)
            #HRA Case Coding - Putting Cases into HRA's Baskets!  
            df['HRA Case Coding'] = df.apply(lambda x: ImmigrationToolBox.HRA_Case_Coding(x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA Level of Service'],x['IOI Does Client Have A Criminal History? (IOI 2)']), axis=1)
            
            #df['Proceeding Type'] = df['Special Legal Problem Code']
                              
        if request.form.get('Employment'):
            print('Employment Version')
            df['HRA Case Coding'] = df.apply(lambda x: EmploymentToolBox.HRA_Case_Coding(x['Level of Service'],x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA IOI Employment Law Retainer?'],x['Matter/Case ID#']), axis = 1) 
            
        df['Proceeding Type'] = df['HRA Case Coding'].apply(lambda x: x[3:] if x != 'Hold For Review' else '')
        
        df['Proceeding Type'] = df['Proceeding Type'].replace('ething is wrong', 'Needs Cleanup')
               
               
        #Sorting Proceeding Type Column
        df = df.sort_values(by=['Proceeding Type'],ascending = False)
        
        #Add Case Number column #8 LS Case number?
        df['Case Number']= "LSNYC"+df['Matter/Case ID#']
              
        #Add Household Size (#) column #9
        df['Household Size (#)'] = df['Number of People 18 and Over']+df['Number of People under 18']
        #.apply(lambda x: int(x[0])) why didn't we need these to do this addition?
        #df['Number of People under 18'].apply(lambda x: int(x[0]))
        
        #Add Household Annual Income ($XX,XXX) column #10
        df['Household Annual Income ($XX,XXX)'] = df['Total Annual Income ']
        
        #Add FPL %	column #11
        df['FPL %'] = df['Percentage of Poverty']
        
        #Add Rent Regulated/NYCHA? column
        #df['Rent Regulated/NYCHA?'] = df['Housing Form Of Regulation'] 
        
        #Add Housing Subsidy? column	
        #df['Housing Subsidy?'] = df['Housing Subsidy Type']
        
        #Add Individual/Group Case?	column #12 unsure
        df['Individual/Group Case?'] = 'Individual'
                      
        #Add Summary of the Request/Other Compelling Factors column	#13
        df['Summary of the Request/Other Compelling Factors'] = ''
        
        #Add Approval? column
        #df['Approval?'] = ''
        
        #Add Date column
        #df['Date'] = ''
        
        #Add Notes/Comments column
        #df['Notes/Comments'] = ''
        
        #Add Empty column in column N/14
        df[''] = ''
        
        #apply hyperlink methodology with splicing and concatenation in column O/15
        def NoIDDelete(CaseID):
            if CaseID == '':
                return 'No Case ID'
            if CaseID.startswith('Unique') == True:
                return 'No Case ID'
            else:
                return str(CaseID)
        df['Matter/Case ID#'] = df.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        
        last7 = df['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = df['Matter/Case ID#']
        df['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del df['Matter/Case ID#']
        move=df['Temp Hyperlinked Case #']
        df.insert(0,'Hyperlinked Case #', move)           
        del df['Temp Hyperlinked Case #']
        
        #add 'Eligible for Waiver Request?' in column p/16 - FPL, prev waiver
        def NeedWaiver(WaiverDate,FPL):
            if pd.isnull(WaiverDate) == False:
                return "No, Waived in on " + str(WaiverDate)
            elif FPL < 201:
                return "No, FPL < 201%"
            else:
                return "Yes"
                               
        if request.form.get('Immigration'):
            df['Eligible for Waiver Request?'] = df.apply(lambda x : NeedWaiver(x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)'],x['Percentage of Poverty']),axis=1)
            
        if request.form.get('Employment'):
            df['Eligible for Waiver Request?'] = df.apply(lambda x: NeedWaiver(x['HRA IOI Employment Law Income Waiver Date'],x['Percentage of Poverty']), axis = 1)
        
        #add Primary advocate in column Q/17
        df['Primary Advocate']= df['Primary Advocate']
        
        #REPORTING VERSION Put everything in the right order
        df = df[['Provider','Contact Person','Date of Request','First & Last Initials','New/Reconsideration?','Income/Tier Upgrade/Other','Proceeding Type','Case Number','Household Size (#)','Household Annual Income ($XX,XXX)','FPL %','Individual/Group Case?','Summary of the Request/Other Compelling Factors','','Hyperlinked Case #','Eligible for Waiver Request?','Primary Advocate']]
              
        """
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)    
        """
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False, header = False,startrow=3)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        #create format for the numbers rows
        numbers_format = workbook.add_format({'text_wrap':True,'bold':True,'valign': 'top','align': 'center','font_name':'Arial Narrow','font_size':9})

        #assign new format to links column 
        worksheet.set_column('O:O',11,link_format)
        
        #Add format for title cell
        TopCell_format = workbook.add_format({'bold':True,'valign': 'top','align': 'left','font_name':'Arial Narrow','font_size':9})
        
        #Add Waiver template header to last column
        worksheet.write('A1', 'Office of Civil Justice - IOI Income Waiver Request Template FY2020',TopCell_format)
        
        #Add income waiver blue/gray header format
        header_format = workbook.add_format({'text_wrap':True,'bold':True,'valign': 'top','align': 'center','bg_color' : '#D6DCE4','font_name':'Arial Narrow','font_size':9})
        
        #Add Yellow header format to columns to remove
        remove_format = workbook.add_format({'text_wrap':True,'bold':True,'valign': 'top','align': 'center','bg_color' : '#FFFF00','font_name':'Arial Narrow','font_size':9})
        #worksheet.set_row(0, None, header_format)
        
        
        worksheet.set_column('F:F',9)
        worksheet.set_column('G:G',17)
        worksheet.set_column('H:H',15)  
        worksheet.set_column('J:J',11)        
        worksheet.set_column('L:L',9)
        worksheet.set_column('M:M',18)
        worksheet.set_column('P:P',25)
        worksheet.set_column('Q:Q',18.5)
        
        #create problem format
        problem_format = workbook.add_format({'bg_color':'yellow'})
        
        #assign problem format to Proceeding Type column
        GRowRange='G4:G'+str(df.shape[0]+3)
        print(GRowRange)
        if request.form.get('Employment'):
            worksheet.conditional_format(GRowRange,{'type': 'cell','criteria': '==','value': '"Needs Cleanup***"','format': problem_format})
        if request.form.get('Immigration'):
            worksheet.conditional_format(GRowRange,{'type': 'cell','criteria': '==','value': '"Needs Cleanup"','format': problem_format})
            
        #Add column headers back in
        for col_num, value in enumerate(df.columns.values):
                    worksheet.write(2, col_num, value, header_format)
        
        #ColNum=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]
        ColNum=list(range(1,14))
        #Add the numbers 1-13 on the column headers        
        for col_num, value in enumerate(ColNum):
                    worksheet.write(1, col_num, value, numbers_format)
                    
        
        
        #Add Borders to everything
        border_format=workbook.add_format({'border':1,'align':'left','font_size':10})
        WaiverRange='A1:W'+str(df.shape[0]+3)
        print(WaiverRange)
        worksheet.conditional_format( WaiverRange, { 'type' : 'cell' ,'criteria': '!=','value':'""','format' : border_format} )
        
        #remove header color from blank column
        worksheet.write('N3','',border_format)
        
        #add different color to column headers for columns to remove before reporting
        worksheet.write('O3','Hyperlinked Case #',remove_format)
        worksheet.write('P3','Eligible for Waiver Request?',remove_format)
        worksheet.write('Q3','Primary Advocate',remove_format)
        
        #worksheet.write('A:W',13,header_format)
        #worksheet.set_row(0, None, header_format)
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Waiver Template " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>IOI Waiver Maker</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Complete Income Waiver Template:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value='Waive In!'>

    </br>
    </br>
    <input type="checkbox" id="Immigration" name="Immigration" value="Immigration">
    <label for="Immigration"> Immigration</label><br>
    <input type="checkbox" id="Employment" name="Employment" value="Employment">
    <label for="Employment"> Employment</label><br>
     
     </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>For immigration requests, check the Immigration box above and use the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1918" target="_blank">Grants Management IOI 2 (3459) Report (Immigration)</a>.</li>
    <li>For employment requests, check the Employment box above and use the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2020" target="_blank">Grants Management IOI Employment (3474) Report</a>.</li>
    </br>
    <a href="/">Home</a>
    '''
    
