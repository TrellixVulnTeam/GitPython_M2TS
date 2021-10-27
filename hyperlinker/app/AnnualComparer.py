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
    
        #reads in four files in arbitrary order
        f = request.files.getlist('file')[0]
        g = request.files.getlist('file')[1]
        h = request.files.getlist('file')[2]
        i = request.files.getlist('file')[3]
        print(f)
        print(g)
        print(h)
        print(i)
        
        #create 4 dataframes from 4 files
        
        df1 = pd.read_excel(f,skiprows=2)
        df2 = pd.read_excel(g,skiprows=2)
        df3 = pd.read_excel(h,skiprows=2)
        df4 = pd.read_excel(i,skiprows=2)
       
       #create an empty list for dataframes and append them
        
        dflist = list()
        dflist.append(df1)
        dflist.append(df2)
        dflist.append(df3)
        dflist.append(df4)
        
        #create empty dataframes to append datasheets to based
        dfopened = pd.DataFrame()
        dfclosed = pd.DataFrame()
        
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
                dfclosed = dfclosed.append(i, ignore_index=True)
        
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

        #consolidate 'close reason' into categories
        dfclosed['Close Reason Category'] = dfclosed.apply(lambda x: DataWizardTools.CloseReasonConsolidator(x['Close Reason']),axis = 1)
        
        #Consolidate various race values into categories for pivot
        dfopened['Race'] = dfopened.apply(lambda x: DataWizardTools.RaceConsolidator(x['Race']),axis=1)
        
        #Make Ages into categories
        dfopened['Client Age at Intake'] = dfopened.apply(lambda x: DataWizardTools.AgeConsolidator(x['Client Age at Intake']),axis=1)
        
        #Make Ages into categories
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
        
        #opened by percentage of poverty
        Poverty_opened_pivot = pd.pivot_table(dfopened,values=['Matter/Case ID#'], index=['Percentage of Poverty'],columns = ['Year Opened'],aggfunc=lambda x: len(x.unique()))
        poverty_reorder = ["0-50%","50-100%","100-200%","200%+"]
        Poverty_opened_pivot = Poverty_opened_pivot.reindex(poverty_reorder)
        print(Poverty_opened_pivot)
        
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        workbook = writer.book
        
        
        
        #Pivot Table and Chart for Cases Closed by Unit
        
        #create new excel tab
        Unit_closed_pivot.to_excel(writer, sheet_name='Unit Closed Pivot')
        #create the formatter for this sheet
        unitclosedpivot = writer.sheets['Unit Closed Pivot']
        #create the chart for this sheet
        unitclosedchart = workbook.add_chart({'type':'column'})
        #set column width
        unitclosedpivot.set_column('A:A',20)
        #get the names from the pivot table
        UnitClosedCategoriesRange="='Unit Closed Pivot'!$A$4:$A$"+str(Unit_closed_pivot.shape[0]+3)
        #get the values for the first & second column from the pivot table
        UnitClosedSeriesOneRange="='Unit Closed Pivot'!$B$4:$B$"+str(Unit_closed_pivot.shape[0]+3)
        UnitClosedSeriesTwoRange="='Unit Closed Pivot'!$C$4:$C$"+str(Unit_closed_pivot.shape[0]+3)
        #Name the y axis
        unitclosedchart.set_y_axis({
            'name': 'Cases Closed',
            'name_font': {'size': 14, 'bold': True},})
        #Name the x axis    
        unitclosedchart.set_x_axis({
            'name': 'Practice Area',
            'name_font': {'size': 14, 'bold': True},}) 
        #set size of chart
        unitclosedchart.set_size({'width': 720, 'height': 576})
        #add pivot table values as 2 series to chart
        unitclosedchart.add_series({
            'categories': UnitClosedCategoriesRange,
            'name': "2020",
            'values': UnitClosedSeriesOneRange})
        unitclosedchart.add_series({
            'name': "2021",
            'values': UnitClosedSeriesTwoRange})
        #hide the useless row
        unitclosedpivot.set_row(0, None, None, {'hidden': True})
        #add chart to the spreadsheet
        unitclosedpivot.insert_chart('D2', unitclosedchart)
        
        
        #Pivot Table and Chart for Cases Opened by Unit
        Unit_opened_pivot.to_excel(writer, sheet_name='Unit Opened Pivot')
        unitopenedpivot = writer.sheets['Unit Opened Pivot']
        unitopenedchart = workbook.add_chart({'type':'column'})
        unitopenedpivot.set_column('A:A',20)
        UnitOpenedCategoriesRange="='Unit Opened Pivot'!$A$4:$A$"+str(Unit_opened_pivot.shape[0]+3)
        UnitOpenedSeriesOneRange="='Unit Opened Pivot'!$B$4:$B$"+str(Unit_opened_pivot.shape[0]+3)
        UnitOpenedSeriesTwoRange="='Unit Opened Pivot'!$C$4:$C$"+str(Unit_opened_pivot.shape[0]+3)
        unitopenedchart.set_y_axis({
            'name': 'Cases Opened',
            'name_font': {'size': 14, 'bold': True},})
        unitopenedchart.set_x_axis({
            'name': 'Practice Area',
            'name_font': {'size': 14, 'bold': True},}) 
        unitopenedchart.set_size({'width': 720, 'height': 576})
        unitopenedchart.add_series({
            'categories': UnitOpenedCategoriesRange,
            'name': "2020",
            'values': UnitOpenedSeriesOneRange})
        unitopenedchart.add_series({
            'name': "2021",
            'values': UnitOpenedSeriesTwoRange})
        unitopenedpivot.set_row(0, None, None, {'hidden': True})
        unitopenedpivot.insert_chart('D2', unitopenedchart)
        
        
        #Pivot Table and Chart for Cases Closed by Close Reason
        Close_reason_pivot.to_excel(writer, sheet_name='Close Reason Pivot')
        closereasonpivot = writer.sheets['Close Reason Pivot']
        closereasonchart = workbook.add_chart({'type':'column'})
        closereasonpivot.set_column('A:A',40)
        CloseReasonCategoriesRange="='Close Reason Pivot'!$A$4:$A$"+str(Close_reason_pivot.shape[0]+3)
        CloseReasonSeriesOneRange="='Close Reason Pivot'!$B$4:$B$"+str(Close_reason_pivot.shape[0]+3)
        CloseReasonSeriesTwoRange="='Close Reason Pivot'!$C$4:$C$"+str(Close_reason_pivot.shape[0]+3)
        closereasonchart.set_y_axis({
            'name': 'Cases Closed',
            'name_font': {'size': 14, 'bold': True},})
        closereasonchart.set_x_axis({
            'name': 'Close Reason',
            'name_font': {'size': 14, 'bold': True},})    
        closereasonchart.set_size({'width': 720, 'height': 576})
        closereasonchart.add_series({
            'categories': CloseReasonCategoriesRange,
            'name': "2020",
            'values': CloseReasonSeriesOneRange})
        closereasonchart.add_series({
            'name': "2021",
            'values': CloseReasonSeriesTwoRange})
        closereasonpivot.set_row(0, None, None, {'hidden': True})
        closereasonpivot.insert_chart('D2', closereasonchart)
        
        #Pivot Table and Chart for Cases Opened by Race
        Race_opened_pivot.to_excel(writer, sheet_name='Opened by Race')
        openracepivot = writer.sheets['Opened by Race']
        openedracechart = workbook.add_chart({'type':'column'})
        openracepivot.set_column('A:A',40)
        OpenedRaceCategoriesRange="='Opened by Race'!$A$4:$A$"+str(Race_opened_pivot.shape[0]+3)
        OpenedRaceSeriesOneRange="='Opened by Race'!$B$4:$B$"+str(Race_opened_pivot.shape[0]+3)
        OpenedRaceSeriesTwoRange="='Opened by Race'!$C$4:$C$"+str(Race_opened_pivot.shape[0]+3)
        openedracechart.set_y_axis({
            'name': 'Cases Opened',
            'name_font': {'size': 14, 'bold': True},})
        openedracechart.set_x_axis({
            'name': 'Client Race',
            'name_font': {'size': 14, 'bold': True},}) 
        openedracechart.set_size({'width': 720, 'height': 576})
        openedracechart.add_series({
            'categories': OpenedRaceCategoriesRange,
            'name': "2020",
            'values': OpenedRaceSeriesOneRange})
        openedracechart.add_series({
            'name': "2021",
            'values': OpenedRaceSeriesTwoRange})
        openracepivot.set_row(0, None, None, {'hidden': True})
        openracepivot.insert_chart('D2', openedracechart)
        
        
        #Pivot Table and Chart for Cases Opened by Age
        Age_opened_pivot.to_excel(writer, sheet_name='Opened by Age')
        openagepivot = writer.sheets['Opened by Age']
        openedagechart = workbook.add_chart({'type':'column'})
        openagepivot.set_column('A:A',40)
        OpenedAgeCategoriesRange="='Opened by Age'!$A$4:$A$"+str(Age_opened_pivot.shape[0]+3)
        OpenedAgeSeriesOneRange="='Opened by Age'!$B$4:$B$"+str(Age_opened_pivot.shape[0]+3)
        OpenedAgeSeriesTwoRange="='Opened by Age'!$C$4:$C$"+str(Age_opened_pivot.shape[0]+3)
        openedagechart.set_y_axis({
            'name': 'Cases Opened',
            'name_font': {'size': 14, 'bold': True},})
        openedagechart.set_x_axis({
            'name': 'Client Age',
            'name_font': {'size': 14, 'bold': True},}) 
        openedagechart.set_size({'width': 720, 'height': 576})
        openedagechart.add_series({
            'categories': OpenedAgeCategoriesRange,
            'name': "2020",
            'values': OpenedAgeSeriesOneRange})
        openedagechart.add_series({
            'name': "2021",
            'values': OpenedAgeSeriesTwoRange})
        openagepivot.set_row(0, None, None, {'hidden': True})
        openagepivot.insert_chart('D2', openedagechart)
        
        #Pivot Table and Chart for Cases Opened by Percentage of Poverty
        Poverty_opened_pivot.to_excel(writer, sheet_name='Opened by Poverty')
        openpovertypivot = writer.sheets['Opened by Poverty']
        openedpovertychart = workbook.add_chart({'type':'column'})
        openpovertypivot.set_column('A:A',20)
        OpenedPovertyCategoriesRange="='Opened by Poverty'!$A$4:$A$"+str(Poverty_opened_pivot.shape[0]+3)
        OpenedPovertySeriesOneRange="='Opened by Poverty'!$B$4:$B$"+str(Poverty_opened_pivot.shape[0]+3)
        OpenedPovertySeriesTwoRange="='Opened by Poverty'!$C$4:$C$"+str(Poverty_opened_pivot.shape[0]+3)
        openedpovertychart.set_y_axis({
            'name': 'Cases Opened',
            'name_font': {'size': 14, 'bold': True},})
        openedpovertychart.set_x_axis({
            'name': 'Client Percentage of Poverty',
            'name_font': {'size': 14, 'bold': True},}) 
        openedpovertychart.set_size({'width': 720, 'height': 576})
        openedpovertychart.add_series({
            'categories': OpenedPovertyCategoriesRange,
            'name': "2020",
            'values': OpenedPovertySeriesOneRange})
        openedpovertychart.add_series({
            'name': "2021",
            'values': OpenedPovertySeriesTwoRange})
        openpovertypivot.set_row(0, None, None, {'hidden': True})
        openpovertypivot.insert_chart('D2', openedpovertychart)
        
        
        
        #raw data tabs
        dfclosed.to_excel(writer, sheet_name='Closed Raw Data',index=False)
        dfopened.to_excel(writer, sheet_name='Opened Raw Data',index=False)
        closedraw = writer.sheets['Closed Raw Data']
        openedraw = writer.sheets['Opened Raw Data']
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        closedraw.set_column('A:A',20,link_format)
        openedraw.set_column('A:A',20,link_format)

        
        #send file back to user
        writer.save()
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Comparison " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Annual Comparer</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Compare FOUR spreadsheets:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file multiple =""><input type=submit value=Compare!>
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
    
