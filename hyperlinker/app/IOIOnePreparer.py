#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/IOIOnePreparer", methods=['GET', 'POST'])
def upload_IOIOnePreparer():
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
                                          
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)  

        
        #Make new column with concatenated names
        df['Complete Client Name']= df['Client First Name']+" "+df['Client Last Name']+" "+df['Date of Birth']
        
        #Turn date opened into numbers
        df['Date Opened Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']),axis=1)
        
        #Turn serviced date into numbers
        df['Service Date Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Service Date']),axis=1)
        
        df['IOI Primary Outcome Date Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['IOI Outcome Date']),axis=1)
        
        df['IOI Secondary Outcome Date Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['IOI Secondary Outcome Date']),axis=1)
        
        df['IOI Tertiary Outcome Date Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['IOI Tertiary Outcome Date']),axis=1)
        
        df=df[df['Date Opened']!=""]
        
        df['Race'] = df.apply(lambda x:DataWizardTools.RaceConsolidator(x['Race']), axis=1)
        
        #Sorting by Date Opened so earliest case is non-duplicate
        df = df.sort_values(by=['Date Opened Construct'],ascending = True)
        
        #Check for case in contract
        def CaseinContract (DOC,SDC):
            if SDC == "":
                SDC = 0          
            if 20220701 <= DOC <= 20230630:
                return "Yes"
            elif 20220701 <= SDC <= 20230630:
                return "Yes"
        
        df['In Contract Year?'] = df.apply(lambda x: CaseinContract(x['Date Opened Construct'],x['Service Date Construct']),axis=1)

        
        df = df[df['In Contract Year?']=="Yes"]
        
        df['bool_series'] = df.duplicated(subset='Complete Client Name')
        #df['Unique Complete Client Name'] = df['Complete Client Name'].unique(), didnt work, lengths don't match
        def UniqueBoolColumn (Bool,CCN):
            if Bool == False:
                return CCN
            else:
                return "DUPLICATE"
        df['U Complete Client Name'] = df.apply(lambda x: UniqueBoolColumn(x['bool_series'],x['Complete Client Name']),axis=1)
        #df['unique items?'] = df['Complete Client Name'].unique()
        print(df['Complete Client Name'].unique())
        print(len(df['Complete Client Name'].unique()))
        
        def EnrollmentQuarterAssigner (Bool, DOC):
            if Bool == True:
                return "None"
            elif 20220701 <= DOC <= 20220930:
                return "Q1"
            elif 20221001 <= DOC <= 20221231:
                return "Q2"
            elif 20230101 <= DOC <= 20230331:
                return "Q3"
            elif 20230401 <= DOC <= 20230630:
                return "Q4"
            else:
                return " Q1 Rollover"
        df['Contract Enrollment Quarter'] = df.apply(lambda x: EnrollmentQuarterAssigner(x['bool_series'],x['Date Opened Construct']),axis=1)
        
        EDF = df[df['Contract Enrollment Quarter']!="None"]
        
        #Tallying filings by quarter (x=x+1 == x+=1)
        def OutcomeQuarterAssigner (PODC,SODC,TODC,QB,QE,PO,SO,TO):
            x=0
            if PODC == "":
                PODC = 0
            if SODC == "":
                SODC = 0
            if TODC == "":
                TODC = 0
            if PO != "":
                if QB <= PODC <= QE:
                    x+=1
            if SO != "":
                if QB <= SODC <= QE:
                    x+=1
            if TO != "":
                if QB <= TODC <= QE:
                    x+=1 
            return x
            
        df['Q1 Outcomes'] = df.apply(lambda x: OutcomeQuarterAssigner(x['IOI Primary Outcome Date Construct'],x['IOI Secondary Outcome Date Construct'],x['IOI Tertiary Outcome Date Construct'],20220701,20220930,x['IOI Outcomes'],x['IOI Secondary Outcomes'],x['IOI Tertiary Outcomes']),axis=1)
        
        df['Q2 Outcomes'] = df.apply(lambda x: OutcomeQuarterAssigner(x['IOI Primary Outcome Date Construct'],x['IOI Secondary Outcome Date Construct'],x['IOI Tertiary Outcome Date Construct'],20221001,20221231,x['IOI Outcomes'],x['IOI Secondary Outcomes'],x['IOI Tertiary Outcomes']),axis=1)
        
        df['Q3 Outcomes'] = df.apply(lambda x: OutcomeQuarterAssigner(x['IOI Primary Outcome Date Construct'],x['IOI Secondary Outcome Date Construct'],x['IOI Tertiary Outcome Date Construct'],20230101,20230331,x['IOI Outcomes'],x['IOI Secondary Outcomes'],x['IOI Tertiary Outcomes']),axis=1)
        
        df['Q4 Outcomes'] = df.apply(lambda x: OutcomeQuarterAssigner(x['IOI Primary Outcome Date Construct'],x['IOI Secondary Outcome Date Construct'],x['IOI Tertiary Outcome Date Construct'],20230401,20230630,x['IOI Outcomes'],x['IOI Secondary Outcomes'],x['IOI Tertiary Outcomes']),axis=1)
            

        df = df[['Hyperlinked CaseID#','bool_series','U Complete Client Name','Contract Enrollment Quarter','Assigned Branch/CC','Primary Assignment','Client Last Name','Client First Name','Complete Client Name','Race','Date of Birth','Date Opened','Date Opened Construct','In Contract Year?','Service Date','Q1 Outcomes','Q2 Outcomes','Q3 Outcomes','Q4 Outcomes','IOI Outcome Date','IOI Outcomes','IOI Secondary Outcome Date','IOI Secondary Outcomes','IOI Tertiary Outcome Date','IOI Tertiary Outcomes']]

        #Construct Summary Tables, EDF unique clients and df includes duplicate clients w multi outcomes
        Enrollment_pivot = pd.pivot_table(EDF,index=['Assigned Branch/CC'],values=['Hyperlinked CaseID#'],columns=['Contract Enrollment Quarter'],aggfunc=lambda x: len(x.unique()), fill_value=0)
        
        Outcomes_pivot = pd.pivot_table(df,index=['Assigned Branch/CC'],values=['Q1 Outcomes','Q2 Outcomes','Q3 Outcomes','Q4 Outcomes'],aggfunc=sum, fill_value=0)
        
        Race_pivot = pd.pivot_table(EDF,index=['Race'],values=['Matter/Case ID#'],aggfunc=lambda x: len(x.unique()), fill_value=0)
        
        #Enrollment_pivot.reset_index(inplace=True)
        print(Enrollment_pivot)
        print(Outcomes_pivot)
        print(Race_pivot)
        
        #A client (with unique birthdate) can only be reported once per contract year (Tally # of unique clients)
        #A client can be counted if case was opened in contract year 
        #or if they have time entered in current contract year (rollovers)
        #A client enrollment is counted in the quarter in which their case was opened
        #All rollovers belong to Q1. Time entered < Date opened
        
        #Filings-any outcome that takes place during the contract year
        #Outcomes have a date that determines their associated quarter
        #Outcomes come from 6 LS columns
        #Client counts once, all outcomes may count
        
        #new pivot table by race
                
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False)
        Enrollment_pivot.to_excel(writer, sheet_name='Pivot Tables',index=True,header = True,startrow=1)
        Outcomes_pivot.to_excel(writer, sheet_name='Pivot Tables',index=True,header = True,startrow=13)
        Race_pivot.to_excel(writer, sheet_name='Pivot Tables',index=True,header = True,startcol=7,startrow=13)
        
        '''PivotList=list(Enrollment_pivot, Outcomes_pivot, Race_pivot)
        for i in Enrollment_pivot:
            print (i)
            tdf = pd.read_excel(i,skiprows=2)
            print(tdf)
       
       #Add dataframes to blank dataframe
        CPdf = df.append(tdf, ignore_index=True)
        print("I am CP "+CPdf)
                      
        
        
        #CPdf = pd.read_excel(writer, sheet_name = 'Pivot Tables', engine = 'xlsxwriter')
        #CPdf = df[Enrollment_pivot] + df[Outcomes_pivot] + df[Race_pivot]
        #print(CPdf)'''
                   
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            #print("I am column width" + str(column_width))
            col_idx = df.columns.get_loc(column)
            #print("I am col idx" + str(col_idx))
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
            
        #column width -  set_column(first_col, last_col, width, cell_format, options)
        for column in Outcomes_pivot:
            Ocolumn_width = (len(column))
            print("I am Ocol wid " + str(Ocolumn_width))
            print("I am col " + str(Outcomes_pivot[column]))
        
        for column in Race_pivot:
            Rcolumn_width = (len(column))
        print("I am Rcol wid " + str(Rcolumn_width))
        print("I am col " + str(Race_pivot[column]))
        
        EHeadings=list(Enrollment_pivot.columns)
        for headinglist,subheading in EHeadings:
            #print(headinglist)
            print(subheading)
            print(len(subheading))
            
        for heading in EHeadings:
            print(heading[1])
            print(len(heading[1]))
        '''for column in Enrollment_pivot['Contract Enrollment Quarter']:
            Ecolumn_width = (len(column))
        print("I am Ecol wid " + str(Ecolumn_width))'''
        #print("I am col " + str(Enrollment_pivot[column]))
        #col_idx = 0
        '''writer.sheets['Pivot Tables'].set_column(col_idx, col_idx, column_width'''
        '''for column in Outcomes_pivot:
            column_width = max(Outcomes_pivot[column].astype(str).map(len).max(), len(column))
            print("I am column " + str(column))
            print("I am col len " + str(len(column)))
            print("I am max " + str(Outcomes_pivot[column].astype(str).map(len).max()))
            #print("I am loc " + str(Outcomes_pivot.loc[1]))
            col_idx = Outcomes_pivot.columns.get_loc(column)
            print("I am col idx" + str(col_idx))'''
        #writer.sheets['Pivot Tables'].set_column(col_idx+1, col_idx+1, column_width)
            
        #print(EDF.iloc[1])
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        #assign new format to column A
        worksheet.set_column('A:A',20,link_format)
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Hyperlinked " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>IOI 1 Preparer</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Prepare Report for IOI 1:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Hyperlink!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2193" target="_blank">"IOI 1 w Outcomes"</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to add case hyperlinks to.</li> 
    <li>Once you have identified this file, click ‘Hyperlink!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    <li>Note, the column with your case ID numbers in it must be titled "Matter/Case ID#" or "id" for this to work.</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''

    