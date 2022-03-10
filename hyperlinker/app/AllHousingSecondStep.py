from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory 
from app import app, db, DataWizardTools, HousingToolBox 
from app.models import User, Post 
from app.forms import PostForm 
from werkzeug.urls import url_parse 
from datetime import datetime 
import pandas as pd 
from zipfile import ZipFile
 
 
@app.route("/AllHousingSecondStep", methods=['GET', 'POST']) 
def AllHousingSecondStep(): 
    if request.method == 'POST': 
        print(request.files['file']) 
        f = request.files['file'] 
         
        test = pd.read_excel(f) 
         
        test.fillna('',inplace=True) 
         
        #Cleaning 
        if test.iloc[0][0] == '': 
            df = pd.read_excel(f,skiprows=2) 
        else: 
            df = pd.read_excel(f) 
         
        print(request.form['SplitCategory'])

        
        #Create Hyperlinks 
        df['Hyperlinked CaseID#'] = df['Hyperlinked CaseID#'].astype(str)
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Hyperlinked CaseID#']),axis=1)                   
        
         
        #sort by case handler 
         
        df = df.sort_values(by=['Primary Advocate']) 
        df = df.sort_values(by=['Assigned Branch/CC']) 
        df = df.sort_values(by=['Tester Tester']) 
         
         
        #Rename second Branch column
        df['Assigned Branch'] = df['Assigned Branch/CC']
        
        #Put everything in the right order 
         
        '''df = df[['Hyperlinked CaseID#', 
        "Assigned Branch", 
        "Tester Tester", 
        'Primary Advocate', 
        "Date Opened", 
        "Date Closed", 
        "Client First Name", 
        "Client Last Name", 
        "Street Address", 
        "Apt#/Suite#", 
        "City", 
        "Zip Code", 
        "HRA Release?", 
        "HAL Eligibility Date", 
        "Pre-12/1/21 Elig Date?", 
        'Unreportable', 
        "Housing Income Verification", 
        #'Income Verification Tester', 
        "Housing Posture of Case on Eligibility Date", 
        "Gen Case Index Number",'Case Number Tester',   
        "Housing Type Of Case", 
        "Housing Level of Service", 
        "Close Reason", 
        "Housing Building Case?", 
        "Primary Funding Code", 
        "Secondary Funding Codes", 
        "Funding Tester", 
        "Housing Total Monthly Rent", 
        #'Rent Tester', 
        "Referral Source", 
        "Gen Pub Assist Case Number",'PA # Tester', 
        "Social Security #", 
        "Housing Number Of Units In Building", 
        "Housing Form Of Regulation", 
        "Housing Subsidy Type", 
        "Housing Years Living In Apartment", 
        "Language", 
        "Housing Activity Indicators", 
        "Housing Services Rendered to Client", 
        "Housing Outcome", 
        "Housing Outcome Date", 
        "Assigned Branch/CC", 
        "Number of People 18 and Over","Number of People under 18","Over-18 Tester", 
        "Date of Birth", 
        "Over 62?", 
        "Percentage of Poverty", 
        "Case Disposition", 
        "Housing Date Of Waiver Approval","Housing TRC HRA Waiver Categories","Waiver Tester", 
        "IOLA Outcome", 
        "Housing Signed DHCI Form", 
        "Income Types", 
        "Total Annual Income ", 
        "Housing Funding Note", 
        "Total Time For Case", 
        "Service Date", 
        "Caseworker Name", 
        "Retainer on File Compliance", 
        "Retainer on File", 
        "Case Involves Covid-19", 
        "Legal Problem Code", 
        "Agency",
        "Housing Tab Assignment",
        #'Post-3/1 Limited Service Tester' 
         
 
        ]] '''      
         
         
        #Preparing Excel Document 
        if request.form['SplitSpreadsheetsBy'] == "Agency":
            print("This is Agency")
            def save_xls(TabSplit,path):
                writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
                #create a dictionary of dataframes for each unique advocate, within the borough that our 'i' for loop is cycling through
                tab_dict_df = dict(tuple(df.groupby(TabSplit)))
                
                #start cycling through each of these new advocate-based dataframes
                for j in tab_dict_df:
                            
                    #write a tab in the borough's excel sheet, composed just of the advocate's cases, with the advocate's name
                    tab_dict_df[j].to_excel(writer, j, index = False)
     
                    workbook = writer.book 
                    ws = writer.sheets[j] 
                    link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True}) 
                    regular_format = workbook.add_format({'font_color':'black'}) 
                    problem_format = workbook.add_format({'bg_color':'yellow'}) 
                    bad_problem_format = workbook.add_format({'bg_color':'red'}) 
                    medium_problem_format = workbook.add_format({'bg_color':'cyan'}) 
                    ws.set_column('A:A',20,link_format) 
                    ws.set_column('B:B',16) 
                    ws.set_column('D:ZZ',25) 
                    ws.set_column('C:C',18) 
                    ws.autofilter('B1:CG1') 
                    ws.freeze_panes(1, 2) 
                    C2BOFullRange='C2:BO'+str(tab_dict_df[j].shape[0]+1) 
                    print("BORowRange is "+ str(C2BOFullRange))
                    CILoc = df.columns.get_loc("Hyperlinked CaseID#")
                    #print(CILoc)
                    #(first_row, first_col, last_row, last_col)
                    shape = (tab_dict_df[j].shape[0]+1)
                    FTLoc=df.columns.get_loc("Funding Tester")
                    #print("0," + str(PALoc) + "," + str(shape) + "," + str(PALoc)   )
                    ws.conditional_format(C2BOFullRange,{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'No Release - Remove Elig Date', 
                                                     'format': bad_problem_format}) 
                    ws.conditional_format(0,FTLoc,shape,FTLoc,{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': "secondary", 
                                                     'format': bad_problem_format}) 
                    ws.conditional_format(0,FTLoc,shape,FTLoc,{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': "Needs", 
                                                     'format': bad_problem_format}) 
                    ws.conditional_format(0,FTLoc,shape,FTLoc,{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': "Targeted", 
                                                     'format': bad_problem_format}) 
                    ws.conditional_format(0,FTLoc,shape,FTLoc,{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': "many", 
                                                     'format': bad_problem_format}) 
                    ws.conditional_format(C2BOFullRange,{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Needs', 
                                                     'format': problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Tester', 
                                                     'format': problem_format}) 
                                                      
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Zip Code', 
                                                     'format': medium_problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Housing Type Of Case', 
                                                     'format': medium_problem_format})                                  
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Gen Case Index Number', 
                                                     'format': medium_problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Language', 
                                                     'format': medium_problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Total Annual Income', 
                                                     'format': medium_problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'HAL Eligibility Date', 
                                                     'format': medium_problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Housing Level of Service', 
                                                     'format': medium_problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Housing Level of Service', 
                                                     'format': medium_problem_format}) 
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Funding Code', 
                                                     'format': medium_problem_format}) 
                                                      
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'HRA Release?', 
                                                     'format': medium_problem_format}) 
                                                      
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Number of People 18 and Over', 
                                                     'format': medium_problem_format})                                                  
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Number of People under 18', 
                                                     'format': medium_problem_format})                                                  
                                                      
                    ws.conditional_format('C1:BO1',{'type': 'text', 
                                                     'criteria': 'containing', 
                                                     'value': 'Date of Birth', 
                                                     'format': medium_problem_format})                                                  
                                                      
                                                      
        
                writer.save()
            output_filename = f.filename
        
            save_xls(TabSplit = request.form['SplitCategory'], path = "app\\sheets\\" + output_filename)
           
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Split " + f.filename)
            
        else:
            def save_xls(ExcelSplit, TabSplit):
                #Create dictionary of dataframes, each named after each unique value of 'assigned branch/cc'
                excel_dict_df = dict(tuple(df.groupby(ExcelSplit)))
                #use 'ZipFile' create empty zip folder, and assign 'newzip' as function-calling name
                with ZipFile("app\\sheets\\zipped\\Split " + f.filename[:-5] +".zip","w") as newzip:
                    
                    #starts cycling through each dataframe (each borough's data)
                    for i in excel_dict_df:
                        
                        #because this is in for loop it creates a new excel file for each 'i' (i.e. each borough)
                        writer = pd.ExcelWriter(path = "app\\sheets\\zipped\\" + i + " " + f.filename[:-5] + ".xlsx", engine = 'xlsxwriter')
                        
                        #create a dictionary of dataframes for each unique advocate, within the borough that our 'i' for loop is cycling through
                        tab_dict_df = dict(tuple(excel_dict_df[i].groupby(TabSplit)))
                        
                        #start cycling through each of these new advocate-based dataframes
                        for j in tab_dict_df:
                            
                            #write a tab in the borough's excel sheet, composed just of the advocate's cases, with the advocate's name
                            tab_dict_df[j].to_excel(writer, j, index = False)
             
                            workbook = writer.book 
                            ws = writer.sheets[j] 
                            link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True}) 
                            regular_format = workbook.add_format({'font_color':'black'}) 
                            problem_format = workbook.add_format({'bg_color':'yellow'}) 
                            bad_problem_format = workbook.add_format({'bg_color':'red'}) 
                            medium_problem_format = workbook.add_format({'bg_color':'cyan'}) 
                            ws.set_column('A:A',20,link_format) 
                            ws.set_column('B:B',16) 
                            ws.set_column('D:ZZ',25) 
                            ws.set_column('C:C',18) 
                            ws.autofilter('B1:CG1') 
                            ws.freeze_panes(1, 2) 
                            C2BOFullRange='C2:BO'+str(tab_dict_df[j].shape[0]+1) 
                            print("BORowRange is "+ str(C2BOFullRange)) 
                            ws.conditional_format(C2BOFullRange,{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'No Release - Remove Elig Date', 
                                                             'format': bad_problem_format}) 
                            ws.conditional_format('AA2:AA100000',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': "secondary", 
                                                             'format': bad_problem_format}) 
                            ws.conditional_format('AA2:AA100000',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': "Needs", 
                                                             'format': bad_problem_format}) 
                            ws.conditional_format('AA2:AA100000',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': "Targeted", 
                                                             'format': bad_problem_format}) 
                            ws.conditional_format('AA2:AA100000',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': "many", 
                                                             'format': bad_problem_format}) 
                            ws.conditional_format(C2BOFullRange,{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Needs', 
                                                             'format': problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Tester', 
                                                             'format': problem_format}) 
                                                              
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Zip Code', 
                                                             'format': medium_problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Housing Type Of Case', 
                                                             'format': medium_problem_format})                                  
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Gen Case Index Number', 
                                                             'format': medium_problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Language', 
                                                             'format': medium_problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Total Annual Income', 
                                                             'format': medium_problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'HAL Eligibility Date', 
                                                             'format': medium_problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Housing Level of Service', 
                                                             'format': medium_problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Housing Level of Service', 
                                                             'format': medium_problem_format}) 
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Funding Code', 
                                                             'format': medium_problem_format}) 
                                                              
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'HRA Release?', 
                                                             'format': medium_problem_format}) 
                                                              
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Number of People 18 and Over', 
                                                             'format': medium_problem_format})                                                  
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Number of People under 18', 
                                                             'format': medium_problem_format})                                                  
                                                              
                            ws.conditional_format('C1:BO1',{'type': 'text', 
                                                             'criteria': 'containing', 
                                                             'value': 'Date of Birth', 
                                                             'format': medium_problem_format})                                                          
                        writer.save()
                        #adds excel file to zipped folder
                        newzip.write("app\\sheets\\zipped\\" + i + " " + f.filename[:-5] + ".xlsx",arcname = i + " " + f.filename[:-5] +  ".xlsx")
         
            output_filename = f.filename[:-5]
            print(output_filename)
             
            #save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename) 
            save_xls(ExcelSplit = request.form['SplitSpreadsheetsBy'], TabSplit = request.form['SplitCategory'])
            
            #return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename) 
            #return send_from_directory('sheets\\zipped','SplitBoroughs.zip', as_attachment = True)
            return send_from_directory('sheets\\zipped','Split ' + f.filename[:-5] +'.zip', as_attachment = True)
 
    return ''' 
    <!doctype html> 
    <title>All Housing Second Step</title> 
    <link rel="stylesheet" href="/static/css/main.css">   
    <h1>All Housing Second Step</h1> 
    <form action="" method=post enctype=multipart/form-data> 
    <p><input type=file name=file><input type=submit value=Split!> 
    
     </br>  
     </br> 
     <label for="SplitSpreadsheetsBy">Choose a split category for each Excel doc:</label>
     <select id="SplitSpreadsheetsBy" name="SplitSpreadsheetsBy">
      <option value="Assigned Branch/CC">Assigned Branch/CC</option>
      <option value="Agency">No Split</option>
      <option value="Tester Tester">Tester Tester</option>
      <option value="Primary Advocate">Primary Advocate</option>
      <option value="Housing Tab Assignment">UAHPLP TRC Other</option>
      
     </select>
     
     </br>
      
    <label for="SplitCategory">Choose a split category for each tab:</label>
    <select id="SplitCategory" name="SplitCategory">
      <option value="Tester Tester">Tester Tester</option>
      <option value="Primary Advocate">Primary Advocate</option>
      <option value="Agency">No Split</option>
      <option value="Assigned Branch/CC">Assigned Branch/CC</option>
      <option value="Housing Tab Assignment">UAHPLP TRC Other</option>
     </select> 
     </br> 

    

    </form> 
    <h3>Instructions:</h3> 
    <ul type="disc"> 
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2452" target="_blank">All Housing Python Report</a>.</li> 
     
     
    </br> 
    <a href="/">Home</a> 
    ''' 
