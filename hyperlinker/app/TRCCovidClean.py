from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd


@app.route("/TRCCovidClean", methods=['GET', 'POST'])
def upload_TRCCovidClean():
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
        
        def HRARelease (HRARelease,EligibilityDate):
            if HRARelease == 'No' and EligibilityDate != '':
                return 'No Release - Remove Elig Date'
            elif HRARelease == '' and EligibilityDate != '':
                return 'No Release - Remove Elig Date'
            elif HRARelease == 'Yes':
                return ''
            else:
                return 'Needs HRA Release'
        df['HRA Release Tester'] = df.apply(lambda x: HRARelease(x['HRA Release?'],x['HAL Eligibility Date']), axis=1)
        
        #Has to have a Housing Type of Case
        
        def HousingType (HousingType):
            if HousingType == '':
                return 'Needs Housing Type of Case'
            else:
                return ''
        df['Housing Type Tester'] = df.apply(lambda x: HousingType(x['Housing Type Of Case']), axis=1)
        
        
        #Has to have a Housing Level of Service - and if DHCR case has to be admin proceeding or rep - admin agency 
        
        
        def HousingLevel (HousingLevel,HousingType):
            if HousingLevel == '':
                return 'Needs Level of Service'
            else:
                return ''
        df['Housing Level Tester'] = df.apply(lambda x: HousingLevel(x['Housing Level of Service'],x['Housing Type Of Case']), axis=1)
        
        #Has to say whether or not it's a building case 
        
        def BuildingCase (BuildingCase):
            if BuildingCase == '':
                return 'Needs Building Case Answer'
            if BuildingCase == 'Prefer Not To Answer':
                return 'Needs Building Case Answer'
            else:
                return ''
        df['Building Case Tester'] = df.apply(lambda x: BuildingCase(x['Housing Building Case?']), axis=1)
        
        #Referral Source Can't Be Blank
        
        def Referral (Referral,FundingSource):
            if Referral == '':
                return 'Needs Referral Source'
            elif FundingSource == '3011 TRC FJC Initiative' and Referral != 'FJC Housing Intake':
                return 'Must be FJC'
            else:
                return ''
        df['Referral Tester'] = df.apply(lambda x: Referral(x['Referral Source'],x['Primary Funding Code']), axis=1)
        
        #monthly rent can't be 0
        
        def Rent (Rent):
            if Rent == 0:
                return 'Needs Rent Amount'
            else:
                return ''
        df['Rent Tester'] = df.apply(lambda x: Rent(x['Housing Total Monthly Rent']), axis=1)
        
        #number of units in building can't be 0
        
        def Units (Units):
            
            Units = str(Units)
            if Units == '0':
                return 'Needs Units'
            elif any(c.isalpha() for c in Units) == True:
                return 'Needs To Be Number'
            else:
                return ''
        df['Unit Tester'] = df.apply(lambda x: Units(x['Housing Number Of Units In Building']), axis=1)
        
        #Housing form of regulation can't be blank
        
        def Regulation (Regulation):
            if Regulation == '':
                return 'Needs Form of Regulation'
            else:
                return ''
        df['Regulation Tester'] = df.apply(lambda x: Regulation(x['Housing Form Of Regulation']), axis=1)
        
        #Housing subsidy can't be blank (can be none)
        
        def Subsidy (Subsidy):
            if Subsidy == '':
                return 'Needs Type of Subsidy'
            else:
                return ''
        df['Subsidy Tester'] = df.apply(lambda x: Subsidy(x['Housing Subsidy Type']), axis=1)
        
        #Years in Apartment Can't be 0 (can be -1)
        
        def Years (Years):
            if Years == 0:
                return 'Needs Years In Apartment'
            elif Years < -1:
                return 'Needs Valid Number'
            else:
                return ''
        df['Years in Apartment Tester'] = df.apply(lambda x: Years(x['Housing Years Living In Apartment']), axis=1)
        
        #Language Can't be Blank
        
        def Language (Language):
            if Language == '':
                return 'Needs Language'
            else:
                return ''
        df['Language Tester'] = df.apply(lambda x: Language(x['Language']), axis=1)
        
        
        #Housing Posture of Case can't be blank if there is an eligibility date
        
        def Posture (Posture,EligibilityDate):
            if EligibilityDate  != '' and Posture == '':
                return 'Needs Posture of Case'
            else:
                return ''
        df['Posture Tester'] = df.apply(lambda x: Posture(x['Housing Posture of Case on Eligibility Date'],x['HAL Eligibility Date']), axis=1)
        
        #Housing Income Verification can't be blank or none and other stuff with kids and poverty level and you just give up if it's closed
        def IncomeVerification (IncomeVerification, Children, PovertyPercent, Disposition):
            if Children > 0 and PovertyPercent <= 200.99 and IncomeVerification == '':
                return 'Must Have DHCI or PA#'
            elif Children > 0 and PovertyPercent <= 200.99 and IncomeVerification == 'None':
                return 'Must Have DHCI or PA#'
            elif Disposition == 'Closed' and IncomeVerification =='None':
                return ''
            elif IncomeVerification == '':
                return 'Needs Income Verification'
            elif IncomeVerification == 'None':
                return 'Needs Income Verification'
            else:
                return ''
        df['Income Verification Tester'] = df.apply(lambda x: IncomeVerification(x['Housing Income Verification'], x['Number of People under 18'], x['Percentage of Poverty'],x['Case Disposition']), axis=1)
       
        #PA Tester (need to be correct format as well)
        def PATester (PANumber):
                        
            PANumber = str(PANumber)
            LastCharacter = PANumber[-1:]
            PenultimateCharater = PANumber[-2:-1]
            SecondCharacter = PANumber [1:2]
            if PANumber == '':
                return ''
            elif PANumber == 'None':
                return ''
            elif SecondCharacter == 'o':
                return ''
            elif SecondCharacter == 'n':
                return ''
            elif len(PANumber) == 10 and str.isalpha(LastCharacter) == True and str.isalpha(PenultimateCharater) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            elif len(PANumber) == 12 and str.isalpha(LastCharacter) == True and str.isalpha(PenultimateCharater) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            elif len(PANumber) == 9 and str.isalpha(LastCharacter) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            else:
                return 'Needs Correct PA # Format'
                
        df['PA # Tester'] = df.apply(lambda x: PATester(x['Gen Pub Assist Case Number']), axis=1)
        
        #Test if case number is correct format (don't need one if it's brief, advice, or out-of-court)
        def CaseNum (CaseNum,Level):
            CaseNum = str(CaseNum)
            First3 = CaseNum[0:3]
            ThirdFromEnd = CaseNum[-3:-2]
            SecondFromEnd = CaseNum[-2:-1]
            First6 = CaseNum[0:6]
            First2 = CaseNum[0:2]
            
            if Level == 'Advice' or Level == 'Brief Service' or Level == 'Out-of-Court Advocacy' or Level == 'Hold For Review':
                return ''            
            #City LT Case format LT-123456-19/XX
            elif len(CaseNum) == 15 and First3 == 'LT-' and ThirdFromEnd == '/':
                return ''
            elif len(CaseNum) == 15 and First3 == 'CV-' and ThirdFromEnd == '/':
                return ''
            #DHCR format AA-123456-S (or 2 letters at end)
            elif str.isalpha(First2) == True and len(CaseNum) == 11 and SecondFromEnd == '-':
                return ''
            elif str.isalpha(First2) == True and len(CaseNum) == 12 and ThirdFromEnd == '-':
                return ''
            #Federal/Supreme format 123456/2019
            elif len(CaseNum) == 11 and str.isdigit(First6) == True:
                return ''
            else:
                return "Needs Correct Case # Format"
                
        df['Case Number Tester'] = df.apply(lambda x: CaseNum(x['Gen Case Index Number'],x['Housing Level of Service']), axis=1)
        
        #Test if social security number is correct format (or ignore it if there's a valid PA number)
        def SSNum (CaseNum, PANumber):
            CaseNum = str(CaseNum)
            First3 = CaseNum[0:3]
            Middle2 = CaseNum[4:6]
            Last4 = CaseNum[7:11]
            FirstDash = CaseNum[3:4]
            SecondDash = CaseNum[6:7]
            PANumber = str(PANumber)
            LastCharacter = PANumber[-1:]
            PenultimateCharater = PANumber[-2:-1]
            SecondCharacter = PANumber [1:2]
            
            if First3 == '000' and Middle2 == '00':
                return 'Needs  Full SS#'
            elif str.isnumeric(First3) == True and str.isnumeric(Middle2) == True and str.isnumeric(Last4) == True and FirstDash == '-' and SecondDash == '-': 
                return ''
            elif len(PANumber) == 10 and str.isalpha(LastCharacter) == True and str.isalpha(PenultimateCharater) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            elif len(PANumber) == 12 and str.isalpha(LastCharacter) == True and str.isalpha(PenultimateCharater) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return ''
            elif len(PANumber) == 9 and str.isalpha(LastCharacter) == False and PenultimateCharater != '-' and PenultimateCharater != ' ':
                return '' 
            else:
                return "Needs Correct SS # Format"
                
        df['SS # Tester'] = df.apply(lambda x: SSNum(x['Social Security #'],x['Gen Pub Assist Case Number']), axis=1)
        
        #Test Housing Activity Indicator - can't be blank for closed cases that are full rep state or full rep federal(housing level of service) and eviction cases(housing type of case: non-payment holdover illegal lockout nycha housing termination)
        
        evictiontypes = ['Holdover','Non-payment','Illegal Lockout','NYCHA Housing Termination']
        leveltypes = ['Representation - State Court','Representation - Federal Court']
        
        def ActivityTester(HousingActivity,Disposition,Level,Type,EvictionTypes,LevelTypes):
            if Disposition == 'Closed' and Level in LevelTypes and Type in EvictionTypes and HousingActivity == '':
                return 'Needs Activity Indicator'
            else:
                return ''
        
       
        df['Housing Activity Tester'] = df.apply(lambda x: ActivityTester(x['Housing Activity Indicators'],x['Case Disposition'],x['Housing Level of Service'],x['Housing Type Of Case'],evictiontypes,leveltypes), axis = 1)
        
        #Test Housing Services Rendered - can't be blank for closed cases that are full rep state or full rep federal(housing level of service)
        
        def ServicesTester(HousingServices,Disposition,Level,Type,EvictionTypes,LevelTypes):
            if Disposition == 'Closed' and Level in LevelTypes and HousingServices == '':
                return 'Needs Services Rendered'
            elif Level == 'Representation - Admin. Agency' and Disposition == 'Closed' and HousingServices == '':
                return 'Needs Services Rendered'
            else:
                return ''
               
        df['Housing Services Tester'] = df.apply(lambda x: ServicesTester(x['Housing Services Rendered to Client'],x['Case Disposition'],x['Housing Level of Service'],x['Housing Type Of Case'],evictiontypes,leveltypes), axis = 1)
        
        #Outcome Tester - needs outcome and date for eviction cases that are full rep at state or federal level (not admin)
        
        def OutcomeTester (Disposition,Outcome,OutcomeDate,Level,Type,EvictionTypes,LevelTypes):
            if Disposition == 'Closed' and Level in LevelTypes and Type in EvictionTypes and Outcome == '' and OutcomeDate == '':
                return 'Needs Outcome & Date'
            elif Disposition == 'Closed' and Level in LevelTypes and Type in EvictionTypes and Outcome == '':
                return 'Needs Outcome'
            elif Disposition == 'Closed' and Level in LevelTypes and Type in EvictionTypes and OutcomeDate == '':
                return 'Needs Outcome Date'
            else:
                return ''
            
        df['Outcome Tester'] = df.apply(lambda x: OutcomeTester(x['Case Disposition'],x['Housing Outcome'],x['Housing Outcome Date'],x['Housing Level of Service'],x['Housing Type Of Case'],evictiontypes,leveltypes), axis = 1)
        
        #COVID Modifications - make the testers blank if it's an advice only pre-3/1 case!
        
        #Differentiate pre- and post- 3/1/20 eligibility date cases
           
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        df['Pre-3/1/20 Elig Date?'] = df.apply(lambda x: HousingToolBox.PreThreeOne(x['DateConstruct']), axis=1)
        
        
        
        
        #CovidException testers to erase clean-up requests
        
        
        df['PA # Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['PA # Tester']), axis=1)
        
        df['SS # Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['SS # Tester']), axis=1)
        
        df['Case Number Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Case Number Tester']), axis=1)
        
        df['Rent Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Rent Tester']), axis=1)
        
        df['Years in Apartment Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Years in Apartment Tester']), axis=1)
        
        df['Referral Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Referral Tester']), axis=1)
        
        df['Income Verification Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Income Verification Tester']), axis=1)
        
        df['Posture Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Posture Tester']), axis=1)
        
        df['Unit Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Unit Tester']), axis=1)
        
        df['Regulation Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Regulation Tester']), axis=1)
        
        df['Subsidy Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Subsidy Tester']), axis=1)
        
        df['Outcome Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Outcome Tester']), axis=1)
        
        df['Housing Services Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Housing Services Tester']), axis=1)
        
        df['Housing Activity Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Housing Activity Tester']), axis=1)
        
        df['HRA Release Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['HRA Release Tester']), axis=1)
        
        df['Housing Type Tester'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['Housing Level of Service'], x['Pre-3/1/20 Elig Date?'], x['Housing Type Tester']), axis=1)
        
      
        
        
        
        
        
        
        
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
        'Pre-3/1/20 Elig Date?'
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
    <title>TRC Cleaner [Covid]</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>TRC Cleanup Report: [Covid]</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1507" target="_blank">TRC Raw Case Data Report</a>.</li>
    
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for TRC cleanup.</li> 
    <li>Once you have identified this file, click ‘Clean!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
