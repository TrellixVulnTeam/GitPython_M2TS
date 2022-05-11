from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/UAHPLPCovidClean", methods=['GET', 'POST'])
def upload_UAHPLPCovidClean():
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
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)          
        
        df['Assigned Branch/CC'] = df.apply(lambda x : DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis=1)   


       
        
        #Has to have an HRA Release
        #add level of service, advice/brief without release and with case index #, NEED release
        def HRAReleaseClean (HRARelease,EligibilityDate,LoS,IndexNum):
            if HRARelease == 'Yes':
                return ''
            elif LoS == 'Advice' or LoS == 'Brief Service': 
                if HRARelease == 'No' or HRARelease == '':
                    if IndexNum != '' and IndexNum.startswith('N') == False and IndexNum.startswith('n') == False:
                        return 'Needs HRA Release due to index num'
            
            if HRARelease == 'No' and EligibilityDate != '':
                return 'No Release - Remove Elig Date'
            elif HRARelease == '' and EligibilityDate != '':
                return 'No Release - Remove Elig Date'
            else:
                return 'Needs HRA Release'
        
        df['HRA Release Tester'] = df.apply(lambda x: HRAReleaseClean(x['HRA Release?'],x['HAL Eligibility Date'],x['Housing Level of Service'],x['Gen Case Index Number']), axis=1)
        
        #Has to have a Housing Type of Case
        
        df['Housing Type Tester'] = df.apply(lambda x: HousingToolBox.HousingTypeClean(x['Housing Type Of Case']), axis=1)
        
        #Has to have a Housing Level of Service 

        df['Housing Level Tester'] = df.apply(lambda x: HousingToolBox.HousingLevelClean(x['Housing Level of Service'],x['Housing Type Of Case']), axis=1)
        
        #Has to say whether or not it's a building case 
      
        df['Building Case Tester'] = df.apply(lambda x: HousingToolBox.BuildingCaseClean(x['Housing Building Case?']), axis=1)
        
        #Referral Source Can't Be Blank
        
        df['Referral Tester'] = df.apply(lambda x: HousingToolBox.ReferralClean(x['Referral Source'],x['Primary Funding Code']), axis=1)
        
        #monthly rent can't be 0
       
        df['Rent Tester'] = df.apply(lambda x: HousingToolBox.RentClean(x['Housing Total Monthly Rent']), axis=1)
        
        #number of units in building can't be 0 or written with letters

        df['Unit Tester'] = df.apply(lambda x: HousingToolBox.UnitsClean(x['Housing Number Of Units In Building']), axis=1)
        
        #Housing form of regulation can't be blank
        df['Regulation Tester'] = df.apply(lambda x: HousingToolBox.RegulationClean(x['Housing Form Of Regulation']), axis=1)
        
        #Housing subsidy can't be blank (can be none)
        df['Subsidy Tester'] = df.apply(lambda x: HousingToolBox.SubsidyClean(x['Housing Subsidy Type']), axis=1)
        
        #Years in Apartment Can't be 0 (can be -1)

        df['Years in Apartment Tester'] = df.apply(lambda x: HousingToolBox.YearsClean(x['Housing Years Living In Apartment']), axis=1)
        
        #Language Can't be Blank
        df['Language Tester'] = df.apply(lambda x: HousingToolBox.LanguageClean(x['Language']), axis=1)
        
        
        #Housing Posture of Case can't be blank if there is an eligibility date
        df['Posture Tester'] = df.apply(lambda x: HousingToolBox.PostureClean(x['Housing Posture of Case on Eligibility Date'],x['HAL Eligibility Date'],x['Housing Type Of Case'],x['Housing Level of Service']), axis=1)
        
        #Housing Income Verification can't be blank or none and other stuff with kids and poverty level and you just give up if it's closed
        
        df['DateClosedConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Closed']), axis=1)
        
        df['Today'] = datetime.now()
        df['TodayConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Today']), axis=1)
        
        df['Income Verification Tester'] = df.apply(lambda x: HousingToolBox.UACIncomeVerificationClean(x['Housing Income Verification'], x['Number of People under 18'], x['Percentage of Poverty'],x['Case Disposition'],x['DateClosedConstruct'],x['TodayConstruct']), axis=1)
       
        #PA Tester (need to be correct format as well)
                
        df['PA # Tester'] = df.apply(lambda x: HousingToolBox.PATesterClean(x['Gen Pub Assist Case Number']), axis=1)
        
        #Test if case number is correct format (don't need one if it's brief, advice, or out-of-court)
        
        df['Case Number Tester'] = df.apply(lambda x: HousingToolBox.CaseNumClean(x['Gen Case Index Number'],x['Housing Level of Service']), axis=1)
        
        #Test if social security number is correct format (or ignore it if there's a valid PA number)
        
                
        df['SS # Tester'] = df.apply(lambda x: HousingToolBox.SSNumClean(x['Social Security #'],x['Gen Pub Assist Case Number']), axis=1)
        
        #Test Housing Activity Indicator - can't be blank for closed cases that are full rep state or full rep federal(housing level of service) and eviction cases(housing type of case: non-payment holdover illegal lockout nycha housing termination)
       
        df['Housing Activity Tester'] = df.apply(lambda x: HousingToolBox.ActivityTesterClean(x['Housing Activity Indicators'],x['Case Disposition'],x['Housing Level of Service'],x['Housing Type Of Case']), axis = 1)
        
        #Test Housing Services Rendered - can't be blank for closed cases that are full rep state or full rep federal(housing level of service)
               
        df['Housing Services Tester'] = df.apply(lambda x: HousingToolBox.ServicesTesterClean(x['Housing Services Rendered to Client'],x['Case Disposition'],x['Housing Level of Service'],x['Housing Type Of Case']), axis = 1)
        
        #Outcome Tester - needs outcome and date for eviction cases that are full rep at state or federal level (not admin) and closed or 6 months since eligibility
        
        df['EligConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        
        df['Outcome Tester'] = df.apply(lambda x: HousingToolBox.UAHPLPOutcomeTesterClean(x['Case Disposition'],x['Housing Outcome'],x['Housing Outcome Date'],x['Housing Level of Service'],x['Housing Type Of Case'],x['EligConstruct'],x['TodayConstruct']), axis = 1)
        
        #COVID Modifications - make the testers blank if it's an advice only pre-3/1 case!
        
        #Differentiate pre- and post- 3/1/20 eligibility date cases
           
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        df['Post 12/1/21 Elig Date?'] = df.apply(lambda x: HousingToolBox.PostTwelveOne(x['DateConstruct']), axis=1)

        #CovidException testers to erase clean-up requests
        #need to know program name and service type, 
        #Translation based on HRA Specs  (this got moved up cuz it's output is necessary for program name)           
        df['proceeding'] = df.apply(lambda x: HousingToolBox.UACProceedingType(x['Housing Type Of Case'],x['Legal Problem Code'],x['Close Reason'],x['Housing Level of Service']), axis=1)
        
        #cases in certain zip codes (RTC zips) that are eviction are UA - everything else is non-UA **Bounce this to housing tools**
        
        def UAorNonUA (TypeOfCase,Zip):
            if TypeOfCase in HousingToolBox.evictionproceedings:
                return "UA"
            else:
                return "Non-UA"
                
        df['program_name'] = df.apply(lambda x: UAorNonUA(x['proceeding'],x['Zip Code']), axis=1)
        #Level of Service becomes Service type 
        df['service_type'] = df.apply(lambda x: HousingToolBox.UACServiceType(x['Housing Level of Service'],x['program_name'],x['Close Reason'],x['Legal Problem Code']), axis=1)

        
        
        df['PA # Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['PA # Tester']), axis=1)
        
        df['SS # Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['SS # Tester']), axis=1)
        
        
        
        
        df['Case Number Tester'] = df.apply(lambda x: HousingToolBox.NoReleaseRedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Case Number Tester'], x['HRA Release?']), axis=1)
        
        df['Rent Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Rent Tester']), axis=1)
        
        df['Years in Apartment Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Years in Apartment Tester']), axis=1)
        
        df['Referral Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Referral Tester']), axis=1)
        
        df['Income Verification Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Income Verification Tester']), axis=1)
        
        df['Posture Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Posture Tester']), axis=1)
        
        df['Unit Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Unit Tester']), axis=1)
        
        df['Regulation Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Regulation Tester']), axis=1)
        
        df['Subsidy Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Subsidy Tester']), axis=1)
        
        df['Outcome Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Outcome Tester']), axis=1)
        
        df['Housing Services Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Housing Services Tester']), axis=1)
        
        df['Housing Activity Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Housing Activity Tester']), axis=1)
        
        df['HRA Release Tester'] = df.apply(lambda x: HousingToolBox.NoReleaseRedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['HRA Release Tester'], x['HRA Release?']), axis=1)
        
        df['Housing Type Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Housing Type Tester']), axis=1)
        
        df['Building Case Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Building Case Tester']), axis=1)
        
    
        #Is everything okay with a case? 

        def TesterTester (ReleaseTester,TypeTester,LevelTester,BuildingTester,ReferralTester,RentTester,UnitTester,RegulationTester,SubsidyTester,YearsTester,LanguageTester,PostureTester,IncomeVerification,PATester,CaseNumberTester,SSTester,ActivityTester,ServicesTester,OutcomeTester,EligibilityDate):
            if ReleaseTester == '' and TypeTester == '' and LevelTester == '' and BuildingTester == '' and ReferralTester == '' and RentTester == '' and UnitTester == '' and RegulationTester == '' and SubsidyTester == '' and YearsTester == '' and LanguageTester == '' and PostureTester == '' and IncomeVerification == '' and PATester == '' and CaseNumberTester == '' and SSTester == '' and ActivityTester == '' and ServicesTester == '' and OutcomeTester == '' and EligibilityDate != '':
                return 'No Cleanup Necessary'
            else:
                return 'Case Needs Attention'
            
        df['Tester Tester'] = df.apply(lambda x: TesterTester(x['HRA Release Tester'],x['Housing Type Tester'],x['Housing Level Tester'],x['Building Case Tester'],x['Referral Tester'],x['Rent Tester'],x['Unit Tester'],x[ 'Regulation Tester'],x['Subsidy Tester'],x['Years in Apartment Tester'],x['Language Tester'],x['Posture Tester'],x['Income Verification Tester'],x['PA # Tester'],x['Case Number Tester'],x['SS # Tester'],x['Housing Activity Tester'],x['Housing Services Tester'],x['Outcome Tester'],x['HAL Eligibility Date']),axis=1)
        

        #sort by case handler
        
        df = df.sort_values(by=['Primary Advocate'])
        df = df.sort_values(by=['Assigned Branch/CC'])
        
        
        #Put everything in the right order
        
        df = df[['Hyperlinked CaseID#','Primary Advocate',
        "Date Opened",
        "Date Closed",
        "Client First Name",
        "Client Last Name",
        "Street Address",
        "City",
        "Zip Code",
        "HRA Release?",'HRA Release Tester',
        "Housing Income Verification",'Income Verification Tester',
        "Gen Case Index Number",'Case Number Tester',        
        "Housing Type Of Case",'Housing Type Tester',
        "Housing Level of Service",'Housing Level Tester',"Close Reason",
        "Housing Building Case?",'Building Case Tester',
        "HAL Eligibility Date","Housing Posture of Case on Eligibility Date",'Posture Tester',
        "Primary Funding Code",
        "Housing Total Monthly Rent",'Rent Tester',
        "Housing Number Of Units In Building",'Unit Tester',
        "Housing Form Of Regulation",'Regulation Tester',
        "Housing Subsidy Type",'Subsidy Tester',
        "Housing Years Living In Apartment",'Years in Apartment Tester',
        "Language",'Language Tester',
        "Gen Pub Assist Case Number",'PA # Tester',
        "Social Security #","SS # Tester",
        "Referral Source",'Referral Tester',
        "Housing Activity Indicators",'Housing Activity Tester',
        "Housing Services Rendered to Client",'Housing Services Tester',
        "Housing Outcome",'Outcome Tester',"Housing Outcome Date",
        "Number of People under 18",
        "Number of People 18 and Over",
        "Percentage of Poverty",
        "Total Annual Income ",
        "Housing Funding Note",
        "Housing Date Of Waiver Approval",
        "Housing TRC HRA Waiver Categories",
        "Date of Birth",
        "Apt#/Suite#","Legal Problem Code","Case Disposition",
        "Assigned Branch/CC",
        "Tester Tester",
        'Post 12/1/21 Elig Date?',
        'program_name',
        'service_type',
        'TodayConstruct',
        'DateClosedConstruct',
        'EligConstruct'
        ]]      
        
        #Preparing Excel Document
        
        #Split into different tabs
        allgood_dictionary = dict(tuple(df.groupby('Tester Tester')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                ws = writer.sheets[i]
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                regular_format = workbook.add_format({'font_color':'black'})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                bad_problem_format = workbook.add_format({'bg_color':'red'})
                medium_problem_format = workbook.add_format({'bg_color':'orange'})
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
                ws.autofilter('B1:ZZ1')
                ws.freeze_panes(1, 2)
                ws.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'No Release - Remove Elig Date',
                                                 'format': bad_problem_format})
                ws.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
                ws.conditional_format('C1:BO1',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Tester',
                                                 'format': problem_format})
                ws.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Must Have DHCI or PA#',
                                                 'format': medium_problem_format})            
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename)
       
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)

    return '''
    <!doctype html>
    <title>UAHPLP Cleaner [Covid]</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>UAHPLP Cleanup Report: [Covid]</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2308" target="_blank">HPLP/UAC Internal Report All Cases</a>.</li>
    
   
    </br>
    <a href="/">Home</a>
    '''
