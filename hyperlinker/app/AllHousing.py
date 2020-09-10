from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/AllHousing", methods=['GET', 'POST'])
def AllHousing():
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
        
        
        
        #Tab Assigner based on Primary Funding Code
        
        def TabAssigner(PrimaryFundingCode):
            if PrimaryFundingCode == "3011 TRC FJC Initiative" or PrimaryFundingCode == "3018 Tenant Rights Coalition (TRC)":
                return "TRC"
            elif PrimaryFundingCode == "3111 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3112 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3113 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3114 HRA-HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3115 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3121 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3122 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3123 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3124 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3125 Universal Access to Counsel – (UAC)":
                return "UAHPLP"
            else:
                return "Other"
                
        df['Funding Code Sorter'] = df.apply(lambda x : TabAssigner(x['Primary Funding Code']),axis=1)
        
        
        df['Agency'] = "LSNYC"

        df['Assigned Branch/CC'] = df.apply(lambda x : DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis=1)   


        #Has to have an HRA Release
        
        df['Release & Elig Tester'] = df.apply(lambda x: HousingToolBox.ReleaseAndEligTester(x['HRA Release?'],x['HAL Eligibility Date']), axis=1)
        
        #Has to have a Housing Type of Case
        
        df['Housing Type Tester'] = df.apply(lambda x: HousingToolBox.HousingTypeClean(x['Housing Type Of Case']), axis=1)
        
        #Has to have a Housing Level of Service 

        df['Housing Level Tester'] = df.apply(lambda x: HousingToolBox.HousingLevelClean(x['Housing Level of Service'],x['Housing Type Of Case']), axis=1)
        
        #Has to say whether or not it's a building case 
      
        df['Building Case Tester'] = df.apply(lambda x: HousingToolBox.BuildingCaseClean(x['Housing Building Case?']), axis=1)
        
        #Referral Source Can't Be Blank
        
        df['Referral Tester'] = df.apply(lambda x: HousingToolBox.ReferralClean(x['Referral Source'],x['Primary Funding Code']), axis=1)
        
        #monthly rent can't be 0 [sharene doesn't care if it's 0]
       
        #df['Rent Tester'] = df.apply(lambda x: HousingToolBox.RentClean(x['Housing Total Monthly Rent']), axis=1)
        df['Rent Tester'] = ''
        
        #number of units in building can't be 0 or written with letters

        df['Unit Tester'] = df.apply(lambda x: HousingToolBox.UnitsClean(x['Housing Number Of Units In Building']), axis=1)
        
        #Housing form of regulation can't be blank
        df['Regulation Tester'] = df.apply(lambda x: HousingToolBox.RegulationClean(x['Housing Form Of Regulation']), axis=1)
        
        #Housing subsidy can't be blank (can be none)
        df['Subsidy Tester'] = df.apply(lambda x: HousingToolBox.SubsidyClean(x['Housing Subsidy Type']), axis=1)
        
        #Years in Apartment Can't be 0 (can be -1)

        df['Years in Apartment Tester'] = df.apply(lambda x: HousingToolBox.YearsClean(x['Housing Years Living In Apartment']), axis=1)
        
        #Language Can't be Blank or Unknown
        df['Language Tester'] = df.apply(lambda x: HousingToolBox.LanguageClean(x['Language']), axis=1)
        
        
        
        
        
        #Housing Income Verification can't be blank or none and other stuff with kids and poverty level and you just give up if it's closed
        
        #df['Income Verification Tester'] = df.apply(lambda x: HousingToolBox.IncomeVerificationClean(x['Housing Income Verification'], x['Number of People under 18'], x['Percentage of Poverty'],x['Case Disposition']), axis=1)
        
        #Test if social security number is correct format (or ignore it if there's a valid PA number)
     
        df['SS # Tester'] = df.apply(lambda x: HousingToolBox.SSNumClean(x['Social Security #'],x['Gen Pub Assist Case Number']), axis=1)
        
        
        #PA Tester (need to be correct format as well)
               
        def PATesterClean (PANumber,SSTester):
                        
            PANumber = str(PANumber)
            LastCharacter = PANumber[-1:]
            PenultimateCharater = PANumber[-2:-1]
            SecondCharacter = PANumber [1:2]
            #if SSTester == '':
            #    return 'Unnecessary due to SS#'
            if PANumber == '' or PANumber == 'None' or PANumber == 'NONE' or SecondCharacter == 'o' or SecondCharacter == 'n':
                return ''
            elif str.isdigit(PANumber) == True and len(PANumber) <= 9:
                return ''
            elif str.isdigit(PANumber[:-1]) == True and str.isalpha(LastCharacter) == True:
                return ''
            elif str.isalpha(PANumber[0:1]) == True and PANumber[1:2] == '-' and str.isdigit(PANumber[2:]) == True:
                return ''
            elif str.isdigit(PANumber[:-2]) == True and PenultimateCharater == '-' and str.isalpha(LastCharacter) == True:
                return ''
            elif len(PANumber) == 10 and str.isalpha(LastCharacter) == True and str.isalpha(PenultimateCharater) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            elif len(PANumber) == 12 and str.isalpha(LastCharacter) == True and str.isalpha(PenultimateCharater) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            elif len(PANumber) == 9 and str.isalpha(LastCharacter) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            else:
                return 'Needs Correct PA # Format'

        
        df['PA # Tester'] = df.apply(lambda x: PATesterClean(x['Gen Pub Assist Case Number'],x['SS # Tester']), axis=1)
        
        #Test if case number is correct format (don't need one if it's brief, advice, or out-of-court)
        
        def CaseNumClean (CaseNum,Level):
            CaseNum = str(CaseNum)
            First3 = CaseNum[0:3]
            Third = CaseNum[2:3]
            ThirdFromEnd = CaseNum[-3:-2]
            SecondFromEnd = CaseNum[-2:-1]
            First6 = CaseNum[0:6]
            First2 = CaseNum[0:2]
            Middle6 = CaseNum[3:9]
            
            if Level == 'Advice' or Level == 'Brief Service' or Level == 'Out-of-Court Advocacy' or Level == 'Hold For Review':
                return '' 
            elif Middle6 == '000000':
                return "Needs Correct Case # Format"
            #City LT Case format LT-123456-19/XX
            elif len(CaseNum) == 15 and First3 == 'LT-' and ThirdFromEnd == '/' and str.isdigit(Middle6) == True:
                return ''
            elif len(CaseNum) == 15 and First3 == 'CV-' and ThirdFromEnd == '/' and str.isdigit(Middle6) == True:
                return ''
            #DHCR format AA-123456-S (or 2 letters at end)
            elif str.isalpha(First2) == True and len(CaseNum) == 11 and SecondFromEnd == '-' and str.isdigit(Middle6) == True and Third == '-':
                return ''
            elif str.isalpha(First2) == True and len(CaseNum) == 12 and ThirdFromEnd == '-' and str.isdigit(Middle6) == True and Third == '-':
                return ''
            #Federal/Supreme format 123456/2019
            elif len(CaseNum) == 11 and str.isdigit(First6) == True:
                return ''
            #NYCHA Housing Termination Cases -  "904607-CR-2019"
            elif len(CaseNum) == 14 and str.isdigit(First6) == True:
                return ''
            else:
                return "Needs Correct Case # Format"
        
        df['Case Number Tester'] = df.apply(lambda x: CaseNumClean(x['Gen Case Index Number'],x['Housing Level of Service']), axis=1)

        
        #Test Housing Services Rendered - can't be blank for closed cases that are full rep state or full rep federal(housing level of service)
         
        def ServicesTesterClean(HousingServices,Disposition,Level,Type):
            
            if Level == 'Representation - Admin. Agency' and HousingServices == '':
                return 'Needs Services Rendered'
            else:
                return ''
         
        df['Housing Services Tester'] = df.apply(lambda x: ServicesTesterClean(x['Housing Services Rendered to Client'],x['Case Disposition'],x['Housing Level of Service'],x['Housing Type Of Case']), axis = 1)
        
        #Housing Type of Case Eviction-Types:
        evictiontypes = ['Holdover','Non-payment','Illegal Lockout']
        #Highest Level of Service Reps
        leveltypes = ['Representation - State Court','Representation - Federal Court']
        
        
        #Calculate client's current age
        def CurrentClientAge(BirthDate):
            if BirthDate == '':
                return 'Needs Date of Birth'
            
            else:
                BirthDateMonth = int(BirthDate[:2])
                BirthDateDay = int(BirthDate[3:5])
                BirthDateYear = int(BirthDate[6:])
                
                Today = datetime.today()
                Today = Today.strftime("%m/%d/%Y")
                
                TodayMonth = int(Today[:2])
                TodayDay = int(Today[3:5])
                TodayYear = int(Today [6:])
                
                #BirthDateConstruct = int(BirthDateYear + BirthDateMonth + BirthDateDay)
                #TodayConstruct = int(TodayYear + TodayMonth + TodayDay)
                
                if TodayYear - BirthDateYear > 62:
                    return "Yes"
                elif TodayYear - BirthDateYear == 62 and TodayMonth > BirthDateMonth:
                    return "Yes"
                elif TodayYear - BirthDateYear == 62 and TodayMonth == BirthDateMonth and BirthDateDay > TodayDay:
                    return "Yes"
                else: 
                    return "No"

            
        df['Over 62?'] = df.apply(lambda x: CurrentClientAge(x['Date of Birth']),axis = 1)
        
        #Test Housing Activity Indicator - can't be blank for cases that are full rep state or full rep federal(housing level of service) and eviction cases(housing type of case: non-payment holdover illegal lockout or nycha housing termination for over-62s)
        
        def ActivityTesterClean(HousingActivity,Over62,Level,Type):
            if Level in leveltypes and Type in evictiontypes and HousingActivity == '':
                return 'Needs Activity Indicator'
            elif Level in leveltypes and Type == "NYCHA Housing Termination" and Over62 == "Yes" and HousingActivity == '':
                return 'Needs Activity Indicator'
            else:
                return ''
        df['Housing Activity Tester'] = df.apply(lambda x: ActivityTesterClean(x['Housing Activity Indicators'],x['Over 62?'],x['Housing Level of Service'],x['Housing Type Of Case']), axis = 1)

        
        #Housing Posture of Case can't be blank if there is an eligibility date
        
        def PostureClean (Posture,EligibilityDate,Type,Level,Over62):
            if Type in evictiontypes and Level.startswith("Rep") == True:
                if EligibilityDate  != '' and Posture == '':
                    return 'Needs Posture of Case'
                else:
                    return ''
            if Type == "NYCHA Housing Termination" and Over62 == "Yes" and Level.startswith("Rep") == True:
                if EligibilityDate  != '' and Posture == '':
                    return 'Needs Posture of Case'
                else:
                    return ''
            else:
                return ''
        
        df['Posture Tester'] = df.apply(lambda x: PostureClean(x['Housing Posture of Case on Eligibility Date'],x['HAL Eligibility Date'],x['Housing Type Of Case'],x['Housing Level of Service'],x['Over 62?']), axis=1)
        
        
        
        #Outcome Tester - needs outcome and date for eviction cases that are full rep at state or federal level (not admin)
        
        def OutcomeTesterClean (Over62,Outcome,OutcomeDate,Level,Type):
            
            if Level in leveltypes and Type in evictiontypes:
                if Outcome == '' and OutcomeDate == '':
                    return 'Needs Outcome & Date'
                elif  Outcome == '':
                    return 'Needs Outcome'
                elif OutcomeDate == '':
                    return 'Needs Outcome Date'
            elif Level in leveltypes and Over62 == "Yes" and Type == "NYCHA Housing Termination":
                if Outcome == '' and OutcomeDate == '':
                    return 'Needs Outcome & Date'
                elif  Outcome == '':
                    return 'Needs Outcome'
                elif OutcomeDate == '':
                    return 'Needs Outcome Date'
            else:
                return ''
        
        df['Outcome Tester'] = df.apply(lambda x: OutcomeTesterClean(x['Over 62?'],x['Housing Outcome'],x['Housing Outcome Date'],x['Housing Level of Service'],x['Housing Type Of Case']), axis = 1)
        

        #Test if Poverty Percentage > 1000%
        
        def PovertyPercentTester (PovertyPercent):
            if PovertyPercent > 1000:
                return "Needs Income Review"
            else:
                return ""
        df['Poverty Percent Tester'] = df.apply(lambda x: PovertyPercentTester(x['Percentage of Poverty']), axis = 1)
        
        #Test # of adults (can't be 0)
        
        def AdultTester (HouseholdAdults):
            if HouseholdAdults == 0:
                return "Needs Over-18 Household Member"
        df['Over-18 Tester'] = df.apply(lambda x: AdultTester(x['Number of People 18 and Over']), axis = 1)
        
        #date of waiver approval & waiver categories - if there's something in one but not the other, then flag it. 
        
        def WaiverTester (WaiverType,WaiverDate):
            if WaiverType != "" and WaiverDate == "":
                return "Needs Waiver Date"
            elif WaiverDate != "" and WaiverType == "":
                return "Needs Waiver Type"
            else:
                return ""
        df['Waiver Tester'] = df.apply(lambda x: WaiverTester(x['Housing TRC HRA Waiver Categories'],x['Housing Date Of Waiver Approval']), axis = 1)
        
        
        #COVID Modifications - make the testers blank if it's an advice only pre-3/1 case!
        
        #Differentiate pre- and post- 3/1/20 eligibility date cases
           
        df['EligConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        df['OpenedConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']), axis=1)
        
        def PreThreeOne(EligibilityDate,OpenedDate):
            if isinstance(EligibilityDate, int) == False:
                if OpenedDate >= 20200701:
                    return "No"
                else:
                    return "Yes"
                
                
                
            elif EligibilityDate < 20200301:
                return "Yes"
            elif EligibilityDate >= 20200301:
                return "No"
        df['Pre-3/1/20 Elig Date?'] = df.apply(lambda x: PreThreeOne(x['EligConstruct'],x['OpenedConstruct']), axis=1)
        
        df['Post-3/1 Limited Service Tester'] = df.apply(lambda x: HousingToolBox.NeedsRedactingTester(x['Housing Level of Service'],x['Pre-3/1/20 Elig Date?'],x['Funding Code Sorter']), axis=1)

        #CovidException testers to erase clean-up requests

        
        df['PA # Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['PA # Tester'],x['Funding Code Sorter']), axis=1)
        
        df['SS # Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['SS # Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Case Number Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Case Number Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Rent Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Rent Tester'],x['Funding Code Sorter']), axis=1)
        
       
        
        df['Years in Apartment Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Years in Apartment Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Referral Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Referral Tester'],x['Funding Code Sorter']), axis=1)
        
        #df['Income Verification Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Income Verification Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Posture Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Posture Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Unit Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Unit Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Regulation Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Regulation Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Subsidy Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Subsidy Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Outcome Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Outcome Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Housing Services Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Housing Services Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Housing Activity Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Housing Activity Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Release & Elig Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Release & Elig Tester'],x['Funding Code Sorter']), axis=1)
        
        df['Housing Type Tester'] = df.apply(lambda x: HousingToolBox.AllHousingRedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Housing Type Tester'],x['Funding Code Sorter']), axis=1)
        

       
        
        #sort by case handler
        
        df = df.sort_values(by=['Primary Advocate'])
        df = df.sort_values(by=['Assigned Branch/CC'])
        
        
        #Put everything in the right order
        
        df = df[['Hyperlinked CaseID#','Primary Advocate',
        'Post-3/1 Limited Service Tester',
        "HRA Release?","HAL Eligibility Date",'Release & Elig Tester',
        "Housing Type Of Case",'Housing Type Tester',
        "Housing Level of Service",'Housing Level Tester',
        "Housing Building Case?",'Building Case Tester',
        "Referral Source",'Referral Tester',
        #'Rent Tester',
        "Housing Number Of Units In Building",'Unit Tester',
        "Housing Form Of Regulation",'Regulation Tester',
        "Housing Subsidy Type",'Subsidy Tester',
        "Housing Years Living In Apartment",'Years in Apartment Tester',
        "Language",'Language Tester',
        "Housing Posture of Case on Eligibility Date",'Posture Tester',
        "Housing Income Verification",
        #'Income Verification Tester',
        "Gen Pub Assist Case Number",'PA # Tester',
        "Gen Case Index Number",'Case Number Tester',  
        "Social Security #","SS # Tester",
        "Housing Activity Indicators",'Housing Activity Tester',
        "Housing Services Rendered to Client",'Housing Services Tester',
        "Housing Outcome",'Outcome Tester',"Housing Outcome Date",
        "Poverty Percent Tester","Percentage of Poverty",
        "Waiver Tester","Housing Date Of Waiver Approval",
        "Housing TRC HRA Waiver Categories",
        
        "Over-18 Tester","Number of People 18 and Over","Number of People under 18",
        
        
        'Pre-3/1/20 Elig Date?',

        "Date Opened",
        "Date Closed",
        "Client First Name",
        
        "Client Last Name",
        #"Street Address",
        "City",
        "Zip Code",
        "Close Reason",
        
        "Primary Funding Code",
        
       
        "Total Annual Income ",
        "Housing Total Monthly Rent",
        "Housing Funding Note",
        
        "Date of Birth",
        "Over 62?",
        #"Apt#/Suite#",
        "Legal Problem Code","Case Disposition",
        "Assigned Branch/CC",
        #"Tester Tester",
        "Agency",

        ]]      
        
        
        #Preparing Excel Document
        
        #Split into different tabs
        allgood_dictionary = dict(tuple(df.groupby('Agency')))
        
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
                ws.set_column('C:C',32)
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
    <title>All Housing</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>All Housing</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1507" target="_blank">TRC Raw Case Data Report</a>.</li>
    
    
    </br>
    <a href="/">Home</a>
    '''
