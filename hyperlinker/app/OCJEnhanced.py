#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/OCJEnhanced", methods=['GET', 'POST'])
def OCJEnhanced():
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
        
        #apply borough-specific advocate-filters
        
        if request.form.get('QLS'):
            df = df[df['Assigned Branch/CC'] == "Queens Legal Services"]
            #everyone?!
            
        elif request.form.get('MLS'):
            df = df[df['Assigned Branch/CC'] == "Manhattan Legal Services"]
            MLS_Advocates = (
            'Abbas, Sayeda',
            'Allen, Sharette',
            'Arboleda, Heather M',
            'Basu, Shantonu J',
            'Black, Rosalind',
            'Braudy, Erica',
            'Chow, Corrin G',
            'Evers, Erin C.',
            'Flores, Aida',
            'Freeman, Daniel A',
            'Frierson, Jerome C',
            'Gokhale, Aparna S',
            'Gonzalez, Matias',
            'Gonzalez, Matias G',
            'Grater, Ashley P',
            'Hao, Lindsay',
            'Harris, Tycel M',
            'Harshberger, Sae',
            'He, Ricky',
            'Mercedes, Jannelys J',
            #'Orsini, Mary K',
            #'Price, Adriana J',
            'Rana, Neil',
            'Risener, Jennifer A',
            'Rockett, Molly C',
            'Lewis Flannory, Hallelujah',
            'Saxton, Jonathan G',
            'Shah, Ami Mahendra',
            'Shamid, Sadia',
            'Sharma, Sagar',
            'Sun, Dao',
            'Surface, Ben L',
            'Tenorio Bocangel, Maricella',
            'Leinbach, August J',
            'MacArthur, Cecilia',
            'Babaturk, Kubra',
            'Basuk, Daniel L',
            'Sugar, Danny',
            'Chu, Katherine B'
            )
            df = df.loc[df['Primary Advocate'].isin(MLS_Advocates)]
            
            
        elif request.form.get('BLS'):
            df = df[df['Assigned Branch/CC'] == "Brooklyn Legal Services"]
            
            BLS_Advocates = (
            #'Ali, Stephanie',
            'Barney, Darryl',
            'Bartholome, Mei Li B.',
            'McLain, Catrina Shanell',
            'Cepeda, Jeanette',
            'Chew, Thomas F',
            'Corsaro, Veronica M',
            'Crowder, Jasmin K',
            'Dolin, Brett A',
            'Drimal, Alex V',
            'Drumm, Kristen E.',
            #'Eisom, Stanley',
            #'Fitzgerald, Mario Q',
            #'Frias De Sosa, Yajaira',
            'Feliz, Alexandra',
            'Gardner III, George C.',
            'Gilfoil, Casey Q',
            'Ginsberg, Irene',
            'Gooding, Nnamdia E',
            #'Groener, Aitan Z',
            #'Haarmann, Landry',
            'Hassan, Ali Hassan Abdelhady',
            'James, Natalie C',
            'Joly, Coco',
            'Kaushal, Mallika',
            'Koepp, Charlie M',
            #'McCammon, Larry A',
            'Lee, Jooyeon',
            'Leroux, Paul A',
            'Loh, Nicholas J',
            'Maltezos, Alexander',
            'McHugh Mills, Maura',
            'Morgan, Dominique A',
            #'Mullen, Evan M',
            'Pepe, Lailah H.',
            'Perez, Amanda M',
            'Pope-Sussman, Raphael A',
            'Record, John-Chris',
            'Reed, Jessica',
            #'Roussos, Katie A',
            'Schiff, Logan J.',
            'Sexton, Andrew R',
            'Stevens, Jean',
            'Sumerall, Iesha R',
            'Sullivan, Jessica T.',
            'Wilson-Wieland, Cherille K',
            'Yavarone, Max J',
            #'Ash, Anissa M',
            #'Harned, Christian C'
            )
            
            df = df.loc[df['Primary Advocate'].isin(BLS_Advocates)]
            """                                                          
            These people aren't in LegalServer advocate list
            Bartholome, Mei Li ***                          
            Feliz, Alexandra   ***     
            Hassan, Ali ***                                      
            Koepp, Charles ****                                     
            Loh, Nicholas ***
            Perez, Amanda ***                   
            Sullivan, Jessica***
            """
            
        elif request.form.get('BxLS'):
            df = df[df['Assigned Branch/CC'] == "Bronx Legal Services"]
            #no list (yet?)
            
        elif request.form.get('SILS'):
            df = df[df['Assigned Branch/CC'] == "Staten Island Legal Services"]
            SILS_Advocates = (
            'Batten, Michael P',
            'Blackburn, Hattie B. Bernice',
            'Falco, Fara M',
            'Martinez, Renee',
            'Sayed, Eman',
            'Winship, Parker L',
            'Puleo Jr, Michael',
            'Nilsson, Erik M',
            'Mulles, Carlos',
            'Loomis, Olivia',
            'Golden, Tashanna B',
            'Fuentes, Maria C')
          
            
            df = df.loc[df['Primary Advocate'].isin(SILS_Advocates)]
            
        else:
            df = df
        
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)    
        
        #determine how to tally each type of case (blank and hold for review = full rep)
        
        def OCJTypeTally(LevelOfService, TypeOfCase):
            if LevelOfService.startswith("A") == True or LevelOfService.startswith("B") == True or LevelOfService.startswith("O") == True:
                return "Advice/Brief/OoC Advocacy"
            elif TypeOfCase.startswith("Hold") == True:
                return "Holdover"
            elif TypeOfCase.startswith("Non") == True:
                return "Non-Pay"
            else:
                return "Other"
        df['OCJ Type Tally'] = df.apply(lambda x : OCJTypeTally(x['Housing Level of Service'],x['Housing Type Of Case']),axis=1)
        
        #determine which full-reps are due to blank/hold-for-review
        
        def OCJSubTally(LevelOfService):
            if LevelOfService == '' or LevelOfService.startswith('Hold') == True:
                return "(Blank LoS/Hold For Review)"
        df['OCJ SubType Tally'] = df.apply(lambda x : OCJSubTally(x['Housing Level of Service']),axis=1)
        
        
        #create new dataframe of just full-rep cases 
        
        FRdf = df[df['Housing Level of Service'] != 'Advice']
        FRdf = FRdf[FRdf['Housing Level of Service'] != 'Brief Service']
        FRdf = FRdf[FRdf['Housing Level of Service'] != 'Out-of-Court Advocacy']
        
        #add column to determine cases that don't have an ERAP status
        
        def NonERAPOther(Stayed,Active,Court):
            if Stayed == "No" or Stayed == "":
                if Active == "No" or Stayed == "":
                    if Court == "No" or Stayed == "":
                        return "Other"
        FRdf['NonERAPOther'] = FRdf.apply(lambda x: NonERAPOther(x['Stayed ERAP Case?'],x['Is stayed ERAP case active?'],x['Paid ERAP Matter restored to Active Court Calendar?']),axis=1)
        
        
        
        #create pivot table for different service types
        advocate_pivot = pd.pivot_table(df,values=['Matter/Case ID#'], index = ['Primary Advocate'], columns=['OCJ Type Tally'], aggfunc = 'count', fill_value = 0)
        
        #make a second pivot table of the sub-type tally
        advocate_pivot2 = pd.pivot_table(df,values=['Matter/Case ID#'], index=['Primary Advocate'], columns=['OCJ SubType Tally'], aggfunc='count', fill_value = 0)
        
        
        
        #mash these pivot tables together
        
        advocate_pivot = pd.concat([advocate_pivot,advocate_pivot2], axis = 1, sort = True)
        advocate_pivot.fillna(0,inplace = True)
        
        #do math on the pivot tables
        advocate_pivot['Matter/Case ID#','Total Cases for OCJ'] = advocate_pivot['Matter/Case ID#','Holdover'] + advocate_pivot['Matter/Case ID#','Non-Pay'] + advocate_pivot['Matter/Case ID#','Other']
        
        advocate_pivot.loc['Total'] = advocate_pivot.sum(axis=0)
        
        
        #split out advocate initials
        advocate_pivot['Matter/Case ID#','Last Initial'] = advocate_pivot.index
        
        advocate_pivot['Matter/Case ID#','First Initial'] = advocate_pivot['Matter/Case ID#','Last Initial'].apply(lambda x: str(x.split(', ')[1:2])[2:3])
        
        advocate_pivot['Matter/Case ID#','Last Initial'] = advocate_pivot['Matter/Case ID#','Last Initial'].apply(lambda x: x[0:1])
        
        
        #create pivot table for second tab giving advocate/case details
        advocate_pivot3 = pd.pivot_table(df,values=['Matter/Case ID#'], index=['Primary Advocate','Primary Funding Code','Housing Type Of Case'], aggfunc='count', fill_value = 0)
        
        
        #make smaller pivot tables to enter onto the first sheet related to ERAP questions
        
        Stayed_ERAP_pivot = pd.pivot_table(FRdf[FRdf['Stayed ERAP Case?'] == 'Yes'],values=['Matter/Case ID#'], index=['Stayed ERAP Case?'], aggfunc='count')
        
        ERAP_Court_pivot = pd.pivot_table(FRdf[FRdf['Paid ERAP Matter restored to Active Court Calendar?'] == 'Yes'],values=['Matter/Case ID#'], index=['Paid ERAP Matter restored to Active Court Calendar?'], aggfunc='count')
        
        Stayed_ERAP_Active_pivot = pd.pivot_table(FRdf[FRdf['Is stayed ERAP case active?'] == 'Yes'],values=['Matter/Case ID#'], index=['Is stayed ERAP case active?'], aggfunc='count')
        
        NonERAPOther_pivot = pd.pivot_table(FRdf,values=['Matter/Case ID#'], index=['NonERAPOther'], aggfunc='count')
        
        #order columns on advocate sheet
        
        '''
        this is an unnecessarily complicated way to rearrange sub-columns in a multi-index dataframe 
        
        multi_tuples = [('Matter/Case ID#','Last Initial'),('Matter/Case ID#','First Initial'),('Matter/Case ID#','(Blank LoS/Hold For Review)'),('Matter/Case ID#','Total Cases for OCJ'), ('Matter/Case ID#','Non-Pay'),('Matter/Case ID#','Holdover'),('Matter/Case ID#','Other')]
        
        multi_cols = pd.MultiIndex.from_tuples(multi_tuples, names=['OCJ Type Tally', 'Primary Advocate'])
        
        advocate_pivot = pd.DataFrame(advocate_pivot, columns=multi_cols)
        '''
        
        
        advocate_pivot = advocate_pivot[[('Matter/Case ID#','Last Initial'),('Matter/Case ID#','First Initial'),('Matter/Case ID#','Total Cases for OCJ'), ('Matter/Case ID#','Non-Pay'),('Matter/Case ID#','Holdover'),('Matter/Case ID#','Other'),('Matter/Case ID#','(Blank LoS/Hold For Review)'),('Matter/Case ID#','Advice/Brief/OoC Advocacy')]]
        
        
        
        #order columns on data sheet
        df['Most Recent Time Entry'] = df['Date of Service']
        
        df = df[['Hyperlinked CaseID#',
        'Primary Advocate',
        'Date Opened',
        'Client First Name',
        'Client Last Name',
        'Housing Type Of Case',
        'Housing Level of Service',
        'Most Recent Time Entry',
        'Case Disposition',
        'Most Recent Note',
        'ERAP Involved Case?',
        'Stayed ERAP Case?',
        'Is stayed ERAP case active?',
        'Paid ERAP Matter restored to Active Court Calendar?',
        'OCJ Type Tally',
        'OCJ SubType Tally'
        ]]
        
        #blank out unnecessary cells
        
        advocate_pivot.ix['Total',0] = ''
        advocate_pivot.ix['Total',1] = ''
        
        advocate_pivot.ix['Total',3] = ''
        advocate_pivot.ix['Total',4] = ''
        advocate_pivot.ix['Total',6] = ''
        advocate_pivot.ix['Total',7] = ''
        advocate_pivot.ix['Total',5] = ''
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        
        
        advocate_pivot.to_excel(writer, sheet_name='ReportingSheet')
        advocate_pivot3.to_excel(writer, sheet_name='AdvocateSheet')
        Stayed_ERAP_pivot.to_excel(writer, sheet_name='ReportingSheet', startcol= 9, startrow = 1)
        ERAP_Court_pivot.to_excel(writer, sheet_name='ReportingSheet', startcol= 9, startrow=4)
        Stayed_ERAP_Active_pivot.to_excel(writer, sheet_name='ReportingSheet', startcol= 9, startrow=7)
        NonERAPOther_pivot.to_excel(writer, sheet_name='ReportingSheet', startcol= 9, startrow=10)
        df.to_excel(writer, sheet_name='DataSheet',index=False)
        #FRdf.to_excel(writer, sheet_name='FRDataSheet',index=False)
        
        
        workbook = writer.book
        worksheet = writer.sheets['DataSheet']
        pivotsheet = writer.sheets['ReportingSheet']
        advocatesheet = writer.sheets['AdvocateSheet']
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        report_format = workbook.add_format({'bg_color':'#82E0AA'})
        fyi_format = workbook.add_format({'bg_color':'#F7DC6F'})
        header_format = workbook.add_format({'bold': True,'valign':'center'})
        border_format=workbook.add_format({'border':1})
        
        
        #assign new format to column A
        worksheet.set_column('A:A',link_format)
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            #print("I am column width" + str(column_width))
            col_idx = df.columns.get_loc(column)
            #print("I am col idx" + str(col_idx))
            writer.sheets['DataSheet'].set_column(col_idx, col_idx, column_width)
        

        
        
        
        worksheet.set_column('A:A',20,link_format)
        
        advocatesheet.set_column('A:C',30)
        advocatesheet.set_column('B:B',70)
        advocatesheet.set_column('D:D',15)

        pivotsheet.set_column('A:A',30)
        pivotsheet.set_column('D:D',20)
        pivotsheet.set_column('H:I',25)
        pivotsheet.set_column('B:C',10)
        pivotsheet.set_column('E:G',10)
        pivotsheet.set_column('J:J',50)
        pivotsheet.set_column('K:K',20)
        
        pivotsheet.write('A3',"Primary Advocate",header_format)
        
        BCRowRange='B4:C'+str(advocate_pivot['Matter/Case ID#'].shape[0]+2)
        
        pivotsheet.conditional_format(BCRowRange,{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': report_format})
        
        DRowRange='D4:D'+str(advocate_pivot['Matter/Case ID#'].shape[0]+3)
        
        pivotsheet.conditional_format(DRowRange,{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': report_format})
        
        EFRowRange='E4:F'+str(advocate_pivot['Matter/Case ID#'].shape[0]+2)
        
        pivotsheet.conditional_format(EFRowRange,{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': report_format})
        
        GIRowRange='G4:I'+str(advocate_pivot['Matter/Case ID#'].shape[0]+2)
        
        pivotsheet.conditional_format(GIRowRange,{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': fyi_format})
                                                 
                                                 
        
        
        
        
        
        
        
        
        
        
       
        
        GRowRange='G4:G'+str(advocate_pivot['Matter/Case ID#'].shape[0]+3)
        
        pivotsheet.conditional_format('K3',{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': report_format})
        pivotsheet.conditional_format('K6',{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': report_format})
        pivotsheet.conditional_format('K9',{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': report_format})
        pivotsheet.conditional_format('K12',{'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 0,
                                                 'format': report_format})
        
        
        #AddBorders
        borderrange='B4:I'+str(advocate_pivot['Matter/Case ID#'].shape[0]+3)
        pivotsheet.conditional_format( borderrange , { 'type' : 'cell' ,
                                                        'criteria': '!=',
                                                        'value':'""',
                                                        'format' : border_format} )
        writer.save()
        
        #send file back to user
        if request.form.get('MLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "MLS " + f.filename)
        elif request.form.get('BLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "BLS " + f.filename)
        elif request.form.get('BxLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "BxLS " + f.filename)
        elif request.form.get('SILS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "SILS " + f.filename)
        elif request.form.get('QLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "QLS " + f.filename)
        else:
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "All Boroughs " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>OCJ Enhanced Report</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Prepare Data for 'Enhanced' OCJ Report:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Prepare!>
    </br>
    </br>
    <input type="checkbox" id="QLS" name="QLS" value="QLS">
    <label for="QLS"> QLS</label><br>
    <input type="checkbox" id="MLS" name="MLS" value="MLS">
    <label for="MLS"> MLS</label><br>
    <input type="checkbox" id="BLS" name="BLS" value="BLS">
    <label for="BLS"> BLS</label><br>
    <input type="checkbox" id="BxLS" name="BxLS" value="BxLS">
    <label for="BxLS"> BxLS</label><br>
    <input type="checkbox" id="SILS" name="SILS" value="SILS">
    <label for="SILS"> SILS</label><br>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    
    This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2514" target="_blank">Enhanced OCJ per-Advocate Data</a>.
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
