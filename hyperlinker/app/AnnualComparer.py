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
    
        #create empty dataframes to append datasheets to base
        dfopened = pd.DataFrame()
        dfclosed = pd.DataFrame()
        
        for cat in request.files.getlist('file'):
            print(cat)
            tdf = pd.read_excel(cat,skiprows=2)
            
            #fill all the blanks with empty strings
            tdf.fillna('',inplace = True)
            #if any of the 'date closed values' are blank, assign to 'opened' dataset
            if '' in tdf['Date Closed'].values:
                print ('Some Cases Still Open')
                tdf['Open/Close Dataset?'] = 'Opened'
                dfopened = dfopened.append(tdf, ignore_index=True)
            #otherwise all the cases are closed and it should go to that sheet
            else:
                print ('All Cases Closed')
                tdf['Open/Close Dataset?'] = 'Closed'
                dfclosed = dfclosed.append(tdf, ignore_index=True)
    
        '''#reads in four files in arbitrary order
        f = request.files.getlist('file')[0]
        g = request.files.getlist('file')[1]
        h = request.files.getlist('file')[2]
        i = request.files.getlist('file')[3]
        j = request.files.getlist('file')[4]
        k = request.files.getlist('file')[5]
        print(f)
        print(g)
        print(h)
        print(i)
        print(j)
        print(k)
        
        #create 4 dataframes from 4 files
        
        df1 = pd.read_excel(f,skiprows=2)
        df2 = pd.read_excel(g,skiprows=2)
        df3 = pd.read_excel(h,skiprows=2)
        df4 = pd.read_excel(i,skiprows=2)
        df5 = pd.read_excel(j,skiprows=2)
        df6 = pd.read_excel(k,skiprows=2)
       
       #create an empty list for dataframes and append them
        
        dflist = list()
        dflist.append(df1)
        dflist.append(df2)
        dflist.append(df3)
        dflist.append(df4)
        dflist.append(df5)
        dflist.append(df6)
        
        
                
        
        #for each data frame in the list
        for i in dflist:
        #fill all the blanks with empty strings
            i.fillna('',inplace = True)
            #if any of the 'date closed values' are blank, assign to 'opened' dataset
            if '' in i['Date Closed'].values:
                print ('Some Cases Still Open')
                i['Open/Close Dataset?'] = 'Opened'
                dfopened = dfopened.append(i, ignore_index=True)
            #otherwise all the cases are closed and it should go to that sheet
            else:
                print ('All Cases Closed')
                i['Open/Close Dataset?'] = 'Closed'
                dfclosed = dfclosed.append(i, ignore_index=True)'''
        
        
        
        #filter dataframes depending on which borough is selected
        
        if not request.form.get('LSU') and not request.form.get('QLS') and not request.form.get('MLS') and not request.form.get('BkLS') and not request.form.get('BxLS') and not request.form.get('SILS'):
            print ("nothing pressed")
        
        else:
            if not request.form.get('LSU'):
                print("LSU not pressed")
                dfopened = dfopened[dfopened['Assigned Branch/CC'] != "Legal Support Unit"]
            if not request.form.get('QLS'):
                print("QLS not pressed")
                dfopened = dfopened[dfopened['Assigned Branch/CC'] != "Queens Legal Services"]
            if not request.form.get('MLS'):
                dfopened = dfopened[dfopened['Assigned Branch/CC'] != "Manhattan Legal Services"]
            if not request.form.get('BkLS'):
                dfopened = dfopened[dfopened['Assigned Branch/CC'] != "Brooklyn Legal Services"]
            if not request.form.get('BxLS'):
                dfopened = dfopened[dfopened['Assigned Branch/CC'] != "Bronx Legal Services"]
            if not request.form.get('SILS'):
                dfopened = dfopened[dfopened['Assigned Branch/CC'] != "Staten Island Legal Services"]


            if not request.form.get('LSU'):
                dfclosed = dfclosed[dfclosed['Assigned Branch/CC'] != "Legal Support Unit"]
            if not request.form.get('QLS'):
                dfclosed = dfclosed[dfclosed['Assigned Branch/CC'] != "Queens Legal Services"]
            if not request.form.get('MLS'):
                dfclosed = dfclosed[dfclosed['Assigned Branch/CC'] != "Manhattan Legal Services"]
            if not request.form.get('BkLS'):
                dfclosed = dfclosed[dfclosed['Assigned Branch/CC'] != "Brooklyn Legal Services"]
            if not request.form.get('BxLS'):
                dfclosed = dfclosed[dfclosed['Assigned Branch/CC'] != "Bronx Legal Services"]
            if not request.form.get('SILS'):
                dfclosed = dfclosed[dfclosed['Assigned Branch/CC'] != "Staten Island Legal Services"]
        
        
        """
        df1.fillna('',inplace = True)
        df2.fillna('',inplace = True)
        df3.fillna('',inplace = True)
        df4.fillna('',inplace = True)
        
        print('df1')
        if '' in df1['Date Closed'].values:
            print ('Some Cases Still Open')
            df1['Open/Close Dataset?'] = 'Opened'
            dfopened = dfopened.append(df1, ignore_index=True)
        else:
            print( 'All Cases Closed')
            df1['Open/Close Dataset?'] = 'Closed'
            dfclosed = dfclosed.append(df1, ignore_index=True)
        
        print('df2')
        if '' in df2['Date Closed'].values:
            print ('Some Cases Still Open')
            df2['Open/Close Dataset?'] = 'Opened'
            dfopened = dfopened.append(df2, ignore_index=True)
        else:
            print( 'All Cases Closed')
            df2['Open/Close Dataset?'] = 'Closed'
            dfclosed = dfclosed.append(df2, ignore_index=True)
        
        print('df3')
        if '' in df3['Date Closed'].values:
            print ('Some Cases Still Open')
            df3['Open/Close Dataset?'] = 'Opened'
            dfopened = dfopened.append(df3, ignore_index=True)
        else:
            print( 'All Cases Closed')
            df3['Open/Close Dataset?'] = 'Closed'
            dfclosed = dfclosed.append(df3, ignore_index=True)
        
        print('df4')
        if '' in df4['Date Closed'].values:
            print ('Some Cases Still Open')
            df4['Open/Close Dataset?'] = 'Opened'
            dfopened = dfopened.append(df4, ignore_index=True)
        else:
            print( 'All Cases Closed')
            df4['Open/Close Dataset?'] = 'Closed'
            dfclosed = dfclosed.append(df4, ignore_index=True)
        """
        
        
        #Create Hyperlinks on raw data tabs
        dfclosed['Matter/Case ID#'] = dfclosed.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)    
        dfopened['Matter/Case ID#'] = dfopened.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1) 
        

        #Separates the year values we'll use for pivot tables
        dfclosed['Year Closed'] = dfclosed['Date Closed'].apply(lambda x: x[-4:] if x !='' else '')
        dfopened['Year Opened'] = dfopened['Date Opened'].apply(lambda x: x[-4:])
        print("Length of date closed "+str(len(dfclosed['Year Closed'].unique())))
        print("Length of date opened "+str(len(dfopened['Year Opened'].unique())))
        print(max(len(dfclosed['Year Closed'].unique()),len(dfopened['Year Opened'].unique())))
        HowManyYears=max(len(dfclosed['Year Closed'].unique()),len(dfopened['Year Opened'].unique()))
        print(HowManyYears)

        #consolidate 'close reason' into categories
        dfclosed['Close Reason Category'] = dfclosed.apply(lambda x: DataWizardTools.CloseReasonConsolidator(x['Close Reason']),axis = 1)
        
        #Consolidate various race values into categories for pivot
        dfopened['Race'] = dfopened.apply(lambda x: DataWizardTools.RaceConsolidator(x['Race']),axis=1)
        
        #Make Ages into categories
        dfopened['Client Age at Intake'] = dfopened.apply(lambda x: DataWizardTools.AgeConsolidator(x['Client Age at Intake']),axis=1)
        
        #Make poverty percentage into categories
        dfopened['Percentage of Poverty'] = dfopened.apply(lambda x: DataWizardTools.PovertyConsolidator(x['Percentage of Poverty']),axis=1)
        
        #Consolidate legal problem codes into units
        dfclosed['Unit'] = dfclosed.apply(lambda x : DataWizardTools.UnitSplitter(x['Legal Problem Code']),axis=1)
        dfopened['Unit'] = dfopened.apply(lambda x : DataWizardTools.UnitSplitter(x['Legal Problem Code']),axis=1)

        
        #Pivot table for cases closed by unit
        Unit_closed_pivot = pd.pivot_table(dfclosed,values=['Matter/Case ID#'], index=['Unit'],columns = ['Year Closed'], aggfunc=lambda x: len(x.unique()))
        
        #Pivot table for cases Opened by unit
        Unit_opened_pivot = pd.pivot_table(dfopened,values=['Matter/Case ID#'], index=['Unit'],columns = ['Year Opened'], aggfunc=lambda x: len(x.unique()))
        
        #pivot table for cases closed by close reason
        
        Close_reason_pivot = pd.pivot_table(dfclosed,values=['Matter/Case ID#'], index=['Close Reason Category'],columns = ['Year Closed'], aggfunc=lambda x: len(x.unique()))
        
        #pivot table for cases opened by race
        Race_opened_pivot = pd.pivot_table(dfopened,values=['Matter/Case ID#'], index=['Race'],columns = ['Year Opened'],aggfunc=lambda x: len(x.unique()))
        
        #opened by age
        Age_opened_pivot = pd.pivot_table(dfopened,values=['Matter/Case ID#'], index=['Client Age at Intake'],columns = ['Year Opened'],aggfunc=lambda x: len(x.unique()))
        
        """
        #Create pivot table for housing cases closed by IOLA Outcome
        dfhousingclosed = dfclosed[dfclosed['Unit'] == 'Housing (Tenant)']
        Outcome_closed_pivot = pd.pivot_table(dfhousingclosed,values=['Matter/Case ID#'], index=['Outcome'],columns = ['Year Closed'],aggfunc=lambda x: len(x.unique()))
        
        print(Outcome_closed_pivot)
        """
        
        #opened by percentage of poverty
        Poverty_opened_pivot = pd.pivot_table(dfopened,values=['Matter/Case ID#'], index=['Percentage of Poverty'],columns = ['Year Opened'],aggfunc=lambda x: len(x.unique()))
        poverty_reorder = ["0-50%","50-100%","100-200%","200%+"]
        Poverty_opened_pivot = Poverty_opened_pivot.reindex(poverty_reorder)
        print(Poverty_opened_pivot)
        
        
        #bounce worksheets back to excel
        output_filename = "YearsCompared.xlsx"     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        workbook = writer.book
        
        
        
        #Pivot Table and Chart for Cases Closed by Unit
        def ChartTabMaker (PivotData, SheetName,YAxis,XAxis):
            #create new excel tab
            PivotData.to_excel(writer, sheet_name=SheetName)
            #create the formatter for this sheet
            TabWriter = writer.sheets[SheetName]
            #create the chart for this sheet
            chart = workbook.add_chart({'type':'column'})
            #set column width
            TabWriter.set_column('A:A',20)
            #Name the y axis
            chart.set_y_axis({
                'name': YAxis,
                'name_font': {'size': 14, 'bold': True},})
            #Name the x axis    
            chart.set_x_axis({
                'name': XAxis,
                'name_font': {'size': 14, 'bold': True},}) 
            #set size of chart
            chart.set_size({'width': 720, 'height': 576})
            #add pivot table values as HowManyYears series to chart
            YearsList=list(range(1,HowManyYears+1))
            for coordinate in (YearsList):
                if HowManyYears >= coordinate:
                    chart.add_series({
                        'categories': [SheetName,3,0,PivotData.shape[0]+2,0],
                        'name': [SheetName,1,coordinate],
                        'values': [SheetName,3,coordinate,PivotData.shape[0]+2,coordinate]})
            '''if HowManyYears >=2:
                chart.add_series({
                    'categories': [SheetName,3,0,PivotData.shape[0]+2,0],
                    'name': [SheetName,1,2],
                    'values': [SheetName,3,2,PivotData.shape[0]+2,2]})
            if HowManyYears >=3:
                chart.add_series({
                    'name': [SheetName,1,3],
                    'values': [SheetName,3,3,PivotData.shape[0]+2,3]})
            if HowManyYears >=4:
                chart.add_series({
                    'name': [SheetName,1,4],
                    'values': [SheetName,3,4,PivotData.shape[0]+2,4]})
            if HowManyYears >=5:
                chart.add_series({
                    'name': [SheetName,1,5],
                    'values': [SheetName,3,5,PivotData.shape[0]+2,5]})'''
            
            #hide the useless row
            TabWriter.set_row(0, None, None, {'hidden': True})
            #add chart to the spreadsheet
            TabWriter.insert_chart(1,HowManyYears+2, chart)
        
        ChartTabMaker(Unit_closed_pivot,'Unit Closed Pivot','Cases Closed','Practice Area')      
        ChartTabMaker(Unit_opened_pivot,'Unit Opened Pivot','Cases Opened','Practice Area')
        ChartTabMaker(PivotData=Close_reason_pivot,SheetName='Close Reason Pivot',YAxis='Cases Closed',XAxis='Close Reason')
        ChartTabMaker(YAxis='Cases Opened',XAxis='Client Race',PivotData=Race_opened_pivot,SheetName='Opened by Race')
        ChartTabMaker(PivotData=Age_opened_pivot,SheetName='Opened by Age',YAxis='Cases Opened',XAxis='Client Age')
        ChartTabMaker(PivotData=Poverty_opened_pivot,SheetName='Opened by Poverty',YAxis='Cases Opened',XAxis='Client Percentage of Poverty')
        
            
        
        #raw data tabs
        dfclosed.to_excel(writer, sheet_name='Closed Raw Data',index=False)
        dfopened.to_excel(writer, sheet_name='Opened Raw Data',index=False)
        closedraw = writer.sheets['Closed Raw Data']
        openedraw = writer.sheets['Opened Raw Data']
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        closedraw.set_column('A:A',20,link_format)
        openedraw.set_column('A:A',20,link_format)

        
        
        
        #send file back to user
        
        
        BoroughsPrefix = ''
        
        if request.form.get('MLS') and request.form.get('BkLS') and request.form.get('BxLS') and request.form.get('SILS') and request.form.get('QLS') and request.form.get('LSU'):
            BoroughsPrefix = 'Citywide '
        elif not request.form.get('MLS') and not request.form.get('BkLS') and not request.form.get('BxLS') and not request.form.get('SILS') and not request.form.get('QLS') and not request.form.get('LSU'):
            BoroughsPrefix = 'Citywide '
        else:
            if request.form.get('MLS'):
                BoroughsPrefix = BoroughsPrefix + 'MLS '
            if request.form.get('BkLS'):
                BoroughsPrefix = BoroughsPrefix + 'BkLS '
            if request.form.get('BxLS'):
                BoroughsPrefix = BoroughsPrefix + 'BxLS '
            if request.form.get('SILS'):
                BoroughsPrefix = BoroughsPrefix + 'SILS '
            if request.form.get('QLS'):
                BoroughsPrefix = BoroughsPrefix + 'QLS '
            if request.form.get('LSU'):
                BoroughsPrefix = BoroughsPrefix + 'LSU '
        
        #unitclosedpivot.write(40,0,"This report contains data for:")
        #unitclosedpivot.write(41,0,BoroughsPrefix)
        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = BoroughsPrefix + output_filename)
     

        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Annual Comparer</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Compare FOUR spreadsheets:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file multiple =""><input type=submit value=Compare!>
    </br>
    </br>
    <input type="checkbox" id="LSU" name="LSU" value="LSU">
    <label for="LSU"> LSU</label><br>
    <input type="checkbox" id="QLS" name="QLS" value="QLS">
    <label for="QLS"> QLS</label><br>
    <input type="checkbox" id="MLS" name="MLS" value="MLS">
    <label for="MLS"> MLS</label><br>
    <input type="checkbox" id="BkLS" name="BkLS" value="BkLS">
    <label for="BkLS"> BkLS</label><br>
    <input type="checkbox" id="BxLS" name="BxLS" value="BxLS">
    <label for="BxLS"> BxLS</label><br>
    <input type="checkbox" id="SILS" name="SILS" value="SILS">
    <label for="SILS"> SILS</label><br>
    
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    
    <li>You will need to upload 4 LegalServer Reports for this tool to function.</li>
    <ul class="square">
    <li>First, run the report: <a href="https://lsnyc.legalserver.org/report/dynamic?load=2424" target="_blank">"Comparison Tool Date Closed"</a> for any given time period this year.</li>
    <li>Second, change the filter of this report to cover the same time period in the previous year with the same date and month.</li>
    <li>Third, run the report: <a href="https://lsnyc.legalserver.org/report/dynamic?load=2425" target="_blank">"Comparison Tool Date Opened"</a> for the same time period you chose above, in the current year.</li>
    <li>Fourth, change the filter of this second report to cover the same time period in the previous year.</li>
    </ul>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
