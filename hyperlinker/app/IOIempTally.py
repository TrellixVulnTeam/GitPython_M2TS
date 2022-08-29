from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, EmploymentToolBox, DataWizardTools

from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import date
import pandas as pd


@app.route("/IOIempTally", methods=['GET', 'POST'])
def upload_IOIempTally():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        test = pd.read_excel(f)
        
        test.fillna('',inplace=True)
        
        #date definition?
        today = date.today()
        print(today.month)
        
        if today.month >= 8:
            howmanymonths = today.month - 7
        else:
            howmanymonths = today.month + 12 - 7
        #howmanymonths = 12
        print(howmanymonths)
       
        #Cleaning
        if test.iloc[0][0] == '':
            df = pd.read_excel(f,skiprows=2)
        else:
            df = pd.read_excel(f)
        
        last7 = df['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = df['Matter/Case ID#']
        df['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        move=df['Temp Hyperlinked Case #']
        df.insert(0,'Hyperlinked Case #', move)           
        del df['Temp Hyperlinked Case #']
        
        df.fillna('',inplace=True)
        
        
        #shorten branch names
        df['Assigned Branch/CC'] = df.apply(lambda x: DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis = 1)
        
        
        #Putting Employment Work in HRA Baskets
        
        df['HRA_Case_Coding'] = df.apply(lambda x: EmploymentToolBox.HRA_Case_Coding(x['Level of Service'],x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA IOI Employment Law Retainer?'],x['Matter/Case ID#']), axis = 1)        

        #Does case need special legal problem code?
                
        df['Special Legal Problem Code'] = df.apply(lambda x: EmploymentToolBox.SPLC_problem(x['Level of Service'],x['Special Legal Problem Code'],x['HRA_Case_Coding']), axis = 1)
        
        
        #Can case be reported based on income?

        df['Exclude due to Income?'] = df.apply(lambda x: EmploymentToolBox.Income_Exclude(x['Percentage of Poverty'],x['HRA IOI Employment Law If Client is Over 200% of FPL, Did you seek a waiver from HRA?']), axis=1)
        
        
        #DateMaker for Date Opened
        
        df['Open Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']),axis = 1)

        
        #DHCI form needed?
        
        
        df['Needs DHCI?'] = df.apply(lambda x: EmploymentToolBox.DHCI_Needed(x['HRA IOI Employment Law DHCI Form?'],x['Level of Service'],x['Open Construct']), axis=1)
        
         
        #Manipulable Dates (this seems like a mess, i would like to fix it later - Jay)            
        
        def Eligibility_Date(Substantial_Activity_Date,Effective_Date,Date_Opened):
            if Substantial_Activity_Date != '':
                return Substantial_Activity_Date
            elif Effective_Date != '':
                return Effective_Date
            else:
                return Date_Opened
        df['Eligibility_Date'] = df.apply(lambda x : Eligibility_Date(x['HRA IOI Employment Law HRA Date Substantial Activity Performed 2023'],x['IOI HRA Effective Date (optional) (IOI 2)'],x['Date Opened']), axis = 1)
        
        
        
        df['Open Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']),axis = 1)
        
        #DateMaker Substantial Activity FY23
        df['Subs Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HRA IOI Employment Law HRA Date Substantial Activity Performed 2023']),axis = 1)
                
        #Substantial Activity for Rollover FY23?

        df['Needs Substantial Activity?'] = df.apply(lambda x: EmploymentToolBox.Needs_Rollover(x['Open Construct'],x['HRA IOI Employment Law HRA Substantial Activity 2023'],x['Subs Construct'],x['Matter/Case ID#']), axis=1)
                     
        #Reportable?
        
        df['Reportable?'] = df.apply(lambda x: EmploymentToolBox.ReportableTester(x['Exclude due to Income?'],x['Needs DHCI?'],x['Needs Substantial Activity?'],x['HRA_Case_Coding']),axis=1)
        
        #Unit of Service Calculator
        df['Units of Service'] = df.apply(lambda x: EmploymentToolBox.UoSCalculator(x['HRA_Case_Coding'],x['Reportable?']),axis=1)
        
        #Assign Outcomes
        def AdviceOutcomeDate(HRAOutcome,HRAOutcomeDate,DateClosed):
            if HRAOutcomeDate == '' and HRAOutcome == 'Advice Given':
                return DateClosed
            else:
                return HRAOutcomeDate
                
        df['HRA IOI Employment Law HRA Outcome Date:'] = df.apply(lambda x: AdviceOutcomeDate(x['HRA IOI Employment Law HRA Outcome:'],x['HRA IOI Employment Law HRA Outcome Date:'],x['Date Closed']), axis = 1)
        
        def AdviceOutcome(HRAOutcome,Employment_Tier,CaseDisposition,HRAOutcomeDate):
            if HRAOutcome == '' and CaseDisposition== 'Closed' and Employment_Tier == 'Advice-No Retainer':
                return 'Advice Given'
            elif HRAOutcome == '' and CaseDisposition == 'Closed':
                return '**Needs Outcome**'
            elif HRAOutcome != '' and CaseDisposition == 'Closed'and HRAOutcomeDate == '':
                return '**Needs Outcome Date**'
            else:
                return HRAOutcome

        df['HRA Outcome'] = df.apply(lambda x: AdviceOutcome(x['HRA IOI Employment Law HRA Outcome:'], x['HRA IOI Employment Law IOI Employment Tier Category:'], x['Case Disposition'],x['HRA IOI Employment Law HRA Outcome Date:']), axis = 1)


        #Better names & HRA Names

        df['Employment Tier Category'] = df['HRA IOI Employment Law IOI Employment Tier Category:']
        
        df['Client Name'] = df['Full Person/Group Name (Last First)']
        
        df['Office'] = df['Assigned Branch/CC']
        
        df['Unique_ID'] = 'LSNYC'+df['Matter/Case ID#']
        
        df['Last_Initial'] = df['Client Last Name'].str[1]
        df['First_Initial'] = df['Client First Name'].str[1]
        
        df['Year_of_Birth'] = df['Date of Birth'].str[-4:]
        
        #gender
        def HRAGender (gender):
            if gender == 'Male' or gender == 'Female':
                return gender
            else:
                return 'Other'
        df['Gender'] = df.apply(lambda x: HRAGender(x['Gender']), axis=1)
        
        df['Country of Origin'] = ''
        
        #county=borough
        df['Borough'] = df['County of Residence']
        
        #household size etc.
        df['Household_Size'] = df['Number of People under 18'].astype(int) + df['Number of People 18 and Over'].astype(int)
        df['Number_of_Children'] = df['Number of People under 18']
        
        #Income Eligible?
        df['Annual_Income'] = df['Total Annual Income ']
        def HRAIncElig (PercentOfPoverty):
            if PercentOfPoverty > 200:
                return 'NO'
            else:
                return 'YES'
        df['Income_Eligible'] = df.apply(lambda x: HRAIncElig(x['Percentage of Poverty']), axis=1)
        
        def IncWaiver (eligible,waiverdate):
            if eligible == 'NO' and waiverdate != '':
                return 'Income'
            else:
                return ''
        df['Waiver_Type'] = df.apply(lambda x: IncWaiver(x['Income_Eligible'],x['HRA IOI Employment Law If Client is Over 200% of FPL, Did you seek a waiver from HRA?']), axis=1)
        
        def IncWaiverDate (waivertype,date):
            if waivertype != '':
                return date
            else:  
                return ''
                
        df['Waiver_Approval_Date'] = df.apply(lambda x: IncWaiverDate(x['Waiver_Type'],x['HRA IOI Employment Law Income Waiver Date']), axis = 1)
        

        #Other Cleanup
        
                        
                 
        def PriorEnrollment (casenumber):
            if casenumber in EmploymentToolBox.ReportedFY20:
                return 'FY 20'
            elif casenumber in EmploymentToolBox.ReportedFY19:
                return 'FY 19'
                
        df['Prior_Enrollment_FY'] = df.apply(lambda x:PriorEnrollment(x['Matter/Case ID#']), axis = 1)
                
              
        df['Service_Type_Code'] = df['HRA_Case_Coding'].apply(lambda x: x[:2] if x != 'Hold For Review' else '')

        df['Proceeding_Type_Code'] = df['HRA_Case_Coding'].apply(lambda x: x[3:] if x != 'Hold For Review' else '')
        
        df['Outcome'] = df['HRA IOI Employment Law HRA Outcome:']
        df['Outcome_Date'] = df['HRA IOI Employment Law HRA Outcome Date:']
        
                             
        
        
        #sorting by borough and advocate
        df = df.sort_values(by=['Office','Primary Advocate'])
        
        
        #Construct Summary Tables
        city_pivot = pd.pivot_table(df,index=['Office'],values=['Units of Service'],aggfunc=sum,fill_value=0)
        
        city_pivot.reset_index(inplace=True)
        
        #remove LSU cases
        city_pivot = city_pivot[city_pivot['Office'] != "LSU"]
        
        #Add Goals to Summary Tables:
                
        def BoroughGoal(Office):
            if Office == "BxLS":
                return 19.06
            elif Office == "BkLS":
                return 38.38
            elif Office == "MLS":
                return 88.26
            elif Office == "QLS":
                return 19.06
            elif Office == "SILS":
                return 6.26
            elif Office == "LSU":
                return 1
            elif Office == "":
                return ""
            else:
                return "hmmmm!"
                              
        city_pivot.reset_index(inplace=True)
        
        #Add goals to City Pivot
               
        city_pivot['Annual Goal'] = city_pivot.apply(lambda x: BoroughGoal(x['Office']), axis=1)
        
        city_pivot['Proportional Goal'] = round((city_pivot['Annual Goal']/12 * howmanymonths),2)
                
                                         
        #sums the goals for each column
        city_pivot.loc['Total','Units of Service':'Proportional Goal'] = city_pivot.sum(axis=0) 
                       
        print (city_pivot) 

        #Add percentage calculator:
        city_pivot['Annual Percentage']=city_pivot['Units of Service']/city_pivot['Annual Goal']
        city_pivot['Proportional Percentage']=city_pivot['Units of Service']/city_pivot['Proportional Goal']        
        
        #REPORTING VERSION Put everything in the right order
        df = df[['Hyperlinked Case #','Office','Primary Advocate','Client Name','Level of Service','Legal Problem Code','Special Legal Problem Code','HRA_Case_Coding','Exclude due to Income?','Needs DHCI?','Needs Substantial Activity?','HRA IOI Employment Law HRA Date Substantial Activity Performed 2023','HRA IOI Employment Law HRA Substantial Activity 2023','Units of Service','Reportable?']]
        city_pivot = city_pivot[['Office','Units of Service','Annual Goal','Annual Percentage','Proportional Goal','Proportional Percentage']]
        
        
        
        
        #Bounce to Excel
        borough_dictionary = dict(tuple(df.groupby('Office')))
           
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                city_pivot.to_excel(writer, sheet_name='City Pivot',index=False,header = False,startrow=1)
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                header_format = workbook.add_format({'text_wrap':True,'bold':True, 'align': 'center', 'bg_color': '#eeece1'})
                PercentHead_format = workbook.add_format({'text_wrap':True,'bold':True,'align': 'center','bg_color': '#d2b48c'})
                percent_format = workbook.add_format({'bold':True})
                percent_format.set_num_format('0.00%')
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                totals_format = workbook.add_format({'bold':True})
                border_format=workbook.add_format({'border':1,'align':'left','font_size':10})
                CityPivot = writer.sheets['City Pivot']
                worksheet = writer.sheets[i]
                worksheet.autofilter('A1:O1')
                worksheet.set_column('A:A',20,link_format)
                worksheet.set_column('B:B',19)
                worksheet.set_column('C:BL',30)
                worksheet.freeze_panes(1,1)
                
                #Add column header data back in
                for col_num, value in enumerate(city_pivot.columns.values):
                    CityPivot.write(0, col_num, value, header_format)
                                  
                #CityPivot.set_row(0,14.5,header_format)
                #CityPivot.set_row(6,20,totals_format)
                CityPivot.set_column('D:D',20,percent_format)
                CityPivot.set_column('F:F',30,percent_format)
                
                CityPivot.write('A7', 'Totals', totals_format)
                CityPivot.write('D1', 'Annual Percentage', PercentHead_format)
                CityPivot.write('F1', 'Proportional Percentage',PercentHead_format)
                CityPivot.set_column('A:C',11.3)
                CityPivot.set_column('D:D',11.3)
                CityPivot.set_column('E:E',11.3)
                CityPivot.set_column('F:F',11.3)
                CityPivot.conditional_format( 'A1:F7' , { 'type' : 'cell' ,'criteria': '!=','value':'""','format' : border_format} )
                
                #Chartmaking Section
                #add chart
                AnnualPercentageChart = workbook.add_chart({'type':'column'})
                AnnualPercentageChart.set_title({'name': 'Progress Toward Annual Goal'})
                #get names from pivot table
                OfficeRange="='City Pivot'!$A$2:$A$6"
                #get the values for the column from the pivot table
                AnnualPercentageRange="='City Pivot'!$D$2:$D$6"
                AnnualGoalsRange="='City Pivot'!$C$2:$C$6"
                #Name the y axis
                AnnualPercentageChart.set_y_axis({
                    'name': 'Percentage of Annual Goal',
                    'name_font': {'size': 14, 'bold': True},
                    'num_format': '0%',
                    'max': 1})
                #Name the x axis    
                AnnualPercentageChart.set_x_axis({
                    'name': 'Office',
                    'name_font': {'size': 14, 'bold': True},}) 
                #set size of chart
                AnnualPercentageChart.set_size({'width': 509, 'height': 413})
                #add pivot table values to chart
                AnnualPercentageChart.add_series({
                    'categories': OfficeRange,
                    'name': "Units Completed",
                    'values': AnnualPercentageRange,
                    'fill': {'color':'#59B3CB'},
                    'border': {'color': 'black'}})
                '''goalscomparisonchart.add_series({
                    'name': "Annual Goals",
                    'values': AnnualGoalsRange,
                    'fill': {'color':'#B3A2C7'},
                    'border': {'color': 'black'}})'''
                
                
                                           
                ERowRange='E1:E'+str(dict_df[i].shape[0]+1)
                print(ERowRange)
                
                worksheet.conditional_format(ERowRange,{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
                worksheet.conditional_format('F1:F100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"***Needs SPLC***"',
                                                 'format': problem_format})
                worksheet.conditional_format('G1:G100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Income Waiver"',
                                                 'format': problem_format})
                worksheet.conditional_format('H1:H100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"***Needs Cleanup***"',
                                                 'format': problem_format})
                worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Income Waiver"',
                                                 'format': problem_format})
                worksheet.conditional_format('J1:J100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs DHCI"',
                                                 'format': problem_format})
                worksheet.conditional_format('K1:K100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Substantial Activity in FY23"',
                                                 'format': problem_format})
                                                 
            #add chart to the spreadsheet
            CityPivot.insert_chart('H1', AnnualPercentageChart)
                                                 
                                   
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)
           

        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleanup " + f.filename)
        
    return '''
    <!doctype html>
    <title>IOI Employment Tally</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Tally your IOI Employment Cases and Case Cleanup:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=IOI-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report (DO) called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2020" target="_blank">"Grants Management IOI Employment (3474) Report"</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for IOI.</li> 
    <li>Once you have identified this file, click ‘IOI-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
