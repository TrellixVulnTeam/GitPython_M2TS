#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
from datetime import date
import pandas as pd
import os



@app.route("/WaiverMaker", methods=['GET', 'POST'])
def WaiverMaker():
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
        
        #Add Provider column
        df['Provider'] = 'LSNYC'
        
        #Add Contact Person column
        df['Contact Person'] = ''
        
        #Add Date of Request column
        df['Date of Request'] = date.today().strftime("%m/%d/%Y")
        
        #Add First & Last Initials column
        df['First & Last Initials'] = df['Client First Name'].apply(lambda x: x[0])+df['Client Last Name'].apply(lambda x: x[0])
                
        #Add New/Reconsideration? column
        df['New/Reconsideration?'] = 'New'
        
        #Add Income/ZIP Code Waiver? column
        df['Income/ZIP Code Waiver?'] = 'Income'
        
        #Add Program (AHTP/UA/Non-UA/HHP) column
        
        def ProgramAssigner(PrimaryFundingCode):
            if PrimaryFundingCode == "3011 TRC FJC Initiative" or PrimaryFundingCode == "3018 Tenant Rights Coalition (TRC)":
                return "AHTP"
            elif PrimaryFundingCode == "3111 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3112 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3113 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3114 HRA-HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3115 HPLP-Homelessness Prevention Law Project":
                return "Non-UA"
            elif PrimaryFundingCode == "3121 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3122 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3123 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3124 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3125 Universal Access to Counsel – (UAC)":
                return "UA"
            else:
                return "Other"
                
        df['Program (AHTP/UA/Non-UA/HHP)'] = df.apply(lambda x : ProgramAssigner(x['Primary Funding Code']),axis=1)
        #What is HHP?
        
        #Add Proceeding Type column
        df['Proceeding Type'] = df['Housing Type Of Case']
        
        #Add L&T Number (if applicable) column
        df['L&T Number (if applicable)']= df['Gen Case Index Number']
        
        #Add ZIP Code (XXXXX) column
        df['ZIP Code (XXXXX)'] = df['Zip Code']
        
        #Add Household Size (#) column
        df['Household Size (#)'] = df['Number of People 18 and Over']+df['Number of People under 18']
        #.apply(lambda x: int(x[0])) why didn't we need these to do this addition?
        #df['Number of People under 18'].apply(lambda x: int(x[0]))
        
        #Add Household Annual Income ($XX,XXX) column
        df['Household Annual Income ($XX,XXX)'] = df['Total Annual Income ']
        
        #Add FPL %	column
        df['FPL %'] = df['Percentage of Poverty']
        
        #df.fillna('',inplace=True)
        
        #Add Rent Regulated/NYCHA? column        
        def RentRegAssigner(RentRegulation):
            if RentRegulation == "Unregulated":
                return "No"
            elif RentRegulation == "Unknown":
                return "Unknown"
            elif pd.isnull(RentRegulation) == True:
                return ''
            else:
                return "Yes"
                
        df['Rent Regulated/NYCHA?'] = df.apply(lambda x : RentRegAssigner(x['Housing Form Of Regulation']),axis=1)
        
                           
        #Add Housing Subsidy? column	
        def SubsidAssigner(SubsidyType):
            if SubsidyType == "None":
                return "None"
            elif pd.isnull(SubsidyType) == True:
                return ''
            else:
                return "Yes"
                
        df['Housing Subsidy?'] = df.apply(lambda x : SubsidAssigner(x['Housing Subsidy Type']),axis=1)
        
        #Add Individual/Group Case?	column
        #df['Individual/Group Case?'] = df['Housing Building Case?']
        
        def IndGrpAssigner(BuildingCase):
            if BuildingCase == "No":
                return "Individual"
            elif BuildingCase == "Yes":
                return "Group"
            else:
                return "Please Check"
                
        df['Individual/Group Case?'] = df.apply(lambda x : IndGrpAssigner(x['Housing Building Case?']),axis=1)
        
        #Add Summary of the Request/Other Compelling Factors column	
        df['Summary of the Request/Other Compelling Factors'] = ''
        
        #Add Approval? column
        df['Approval?'] = ''
        
        #Add Date column
        df['Date'] = ''
        
        #Add Notes/Comments column
        df['Notes/Comments'] = ''
        
        '''#Add Already Waived In? column - hra referrals and eviction w court case
        def WaivedIn(Ref, FPL, CaseType, CaseNum):
            CaseNum = str(CaseNum)
            CaseType = str(CaseType)
            if Ref == "HRA" and FPL >= 201:
                return "HRA Referral"
            elif CaseType == "Holdover" or "Illegal Lockout" or "Non-payment" or "NYCHA Housing Termination":
                return "EVC w Court Case"
            else:
                return "No"'''
                
        #Add Already Waived In? column - hra referrals and eviction w court case
        def WaivedIn(Ref,FPL,CaseType,CaseNum):
            CaseType = str(CaseType)
            CaseNum = str(CaseNum)
            #CaseNum = CaseNum.lower()
            if Ref == "HRA" and FPL >= 201:
                return "HRA Referral"
            elif FPL < 201:
                return "No Need"
            elif CaseType in HousingToolBox.evictiontypes:
                if CaseNum.startswith('n') == False and CaseNum.startswith('N') == False: 
                    return "EVC w Court Case"
                else:
                    return "No"
            else:
                return "No"
                
        df['Already Waived In?'] = df.apply(lambda x : WaivedIn(x['Referral Source'],x['Percentage of Poverty'],x['Housing Type Of Case'],x['Gen Case Index Number']),axis=1)
        
        '''CaseType == "Holdover"or"Non-payment"or"Illegal Lockout"or"NYCHA Housing Termination"and'''
        '''and(CaseNum.startswith('N')or CaseNum= == False'''
        '''and CaseType == "Holdover" or CaseType == "Non-payment" or CaseType == "Illegal Lockout" or CaseType == "NYCHA Housing Termination"'''
        
        #Add Eligible for Waiver Request? column
        def NeedWaiver(WaiverType,Ref,FPL,Waived):
            if pd.isnull(WaiverType) == False:
                return WaiverType
            elif FPL < 201:
                return "FPL < 201%"
            elif Ref == "HRA" and FPL >= 201:
                return "HRA Referral"
            elif Waived ==  "EVC w Court Case":
                return "Categorically Waived In" 
            else:
                return "Yes"
                
        df['Eligible for Waiver Request?'] = df.apply(lambda x : NeedWaiver(x['Housing TRC HRA Waiver Categories'],x['Referral Source'],x['Percentage of Poverty'],x['Already Waived In?']),axis=1)
        
        #Sorting Eligible Cases on top
        df = df.sort_values(by=['Eligible for Waiver Request?'],ascending = False)
        
        
        #Add Empty column
        df[''] = ''
        
        #REPORTING VERSION Put everything in the right order
        df = df[['Provider','Contact Person','Date of Request','First & Last Initials','New/Reconsideration?','Income/ZIP Code Waiver?','Program (AHTP/UA/Non-UA/HHP)','Proceeding Type','L&T Number (if applicable)','ZIP Code (XXXXX)','Household Size (#)','Household Annual Income ($XX,XXX)','FPL %','Rent Regulated/NYCHA?','Housing Subsidy?','Individual/Group Case?','Summary of the Request/Other Compelling Factors','Approval?','Date','Notes/Comments','','Hyperlinked Case #','Already Waived In?','Eligible for Waiver Request?']]
              
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
        
        #create problem format
        problem_format = workbook.add_format({'bg_color':'yellow'})

        #assign new format to links column 
        worksheet.set_column('V:V',11,link_format)
        
        #Define ranges so problem format doesn't apply forever
        PRowRange='P4:P'+str(df.shape[0]+3)
        print(PRowRange)
        
        #assign problem format to ind/group case column
        worksheet.conditional_format(PRowRange,{'type': 'cell','criteria': '==','value': '"please check"','format': problem_format})
                 
        
        #Add format for title cell
        TopCell_format = workbook.add_format({'bold':True,'valign': 'top','align': 'left','font_name':'Arial Narrow','font_size':9})
        
        #Add Waiver template header to last column
        worksheet.write('A1', 'Office of Civil Justice - Housing Income/ZIP Code Waiver Request Template FY2020',TopCell_format)
        
        #Add income waiver blue/gray header format
        header_format = workbook.add_format({'text_wrap':True,'bold':True,'valign': 'top','align': 'center','bg_color' : '#D6DCE4','font_name':'Arial Narrow','font_size':9})
        #worksheet.set_row(0, None, header_format)
        
        
        worksheet.set_column('G:G',17)        
        worksheet.set_column('L:L',19)
        worksheet.set_column('Q:Q',26)
       
        #Add column headers back in
        for col_num, value in enumerate(df.columns.values):
                    worksheet.write(2, col_num, value, header_format)
        
        #ColNum=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]
        ColNum=list(range(1,21))
        #Add the numbers 1-20 on the column headers        
        for col_num, value in enumerate(ColNum):
                    worksheet.write(1, col_num, value, numbers_format)
                    
        #Add Borders to everything
        border_format=workbook.add_format({'border':1,'align':'left','font_size':10})
        WaiverRange='A1:X'+str(df.shape[0]+3)
        print(WaiverRange)
        worksheet.conditional_format( WaiverRange, { 'type' : 'cell' ,'criteria': '!=','value':'""','format' : border_format} )
        
        #remove header color from blank column
        worksheet.write('U3','',border_format)
        
        #worksheet.write('A:W',13,header_format)
        #worksheet.set_row(0, None, header_format)
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Waiver Template " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Waiver Maker</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Complete Income Waiver Template:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value='Waive In!'>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1507" target="_blank">TRC Raw Case Data Report</a>.</li>
    </br>
    <a href="/">Home</a>
    '''
    
