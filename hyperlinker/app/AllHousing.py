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
        def UnitsClean (Units):
            if Units == 0:
                return 'Needs Units'
            elif Units < -1:
                return "Needs Valid Unit #"
            Units = str(Units)
            if any(c.isalpha() for c in Units) == True:
                return 'Needs To Be Number'
            elif Units == "0":
                return 'Needs Units'
            else:
                return ''
        df['Unit Tester'] = df.apply(lambda x: UnitsClean(x['Housing Number Of Units In Building']), axis=1)
        
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
        
        def SSNumClean (CaseNum):
            CaseNum = str(CaseNum)
            First3 = CaseNum[0:3]
            Middle2 = CaseNum[4:6]
            Last4 = CaseNum[7:11]
            FirstDash = CaseNum[3:4]
            SecondDash = CaseNum[6:7]

            
            if First3 == '000' and Middle2 == '00':
                return 'Needs  Full SS#'
            elif str.isnumeric(First3) == True and str.isnumeric(Middle2) == True and str.isnumeric(Last4) == True and FirstDash == '-' and SecondDash == '-': 
                return ''
            else:
                return "Needs Correct SS # Format"
                
        df['SS # Tester'] = df.apply(lambda x: SSNumClean(x['Social Security #']), axis=1)
        
        
        #PA Tester (need to be correct format as well)
               
        def PATesterClean (PANumber,SSTester):
                        
            PANumber = str(PANumber)
            LastCharacter = PANumber[-1:]
            PenultimateCharater = PANumber[-2:-1]
            SecondCharacter = PANumber [1:2]
            if SSTester == '':
                return 'Unnecessary due to SS#'
            elif PANumber == '' or PANumber == 'None' or PANumber == 'NONE' or SecondCharacter == 'o' or SecondCharacter == 'n':
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
        
        #if PA # Tester is fine then SS# Tester doesn't matter
        
        def SSDoubleTester(SSTester, PATester, PANum):
            PANum = str(PANum)
            SecondCharacter = PANum [1:2]
            if PANum == '' or PANum == 'None' or PANum == 'NONE' or SecondCharacter == 'o' or SecondCharacter == 'n':
                return SSTester
            elif PATester == '' and len(PANum) >= 9:
                return 'Unnecessary due to PA #'
            else:
                return SSTester
        
        df['SS # Tester'] = df.apply(lambda x: SSDoubleTester(x['SS # Tester'],x['PA # Tester'],x['Gen Pub Assist Case Number']),axis=1)
        
        
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
                
                if TodayMonth == 1:
                    TodayMonth = 13
                
                #BirthDateConstruct = int(BirthDateYear + BirthDateMonth + BirthDateDay)
                #TodayConstruct = int(TodayYear + TodayMonth + TodayDay)
                
                if TodayYear - BirthDateYear > 62:
                    return "Yes"
                elif TodayYear - BirthDateYear == 62 and TodayMonth - 1 >= BirthDateMonth:
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
        
        def OutcomeTesterClean (Over62,Outcome,OutcomeDate,Level,Type,Disposition):
            
            if Level in leveltypes and Type in evictiontypes and Disposition == "Closed":
                if Outcome == '' and OutcomeDate == '':
                    return 'Needs Outcome & Date'
                elif  Outcome == '':
                    return 'Needs Outcome'
                elif OutcomeDate == '':
                    return 'Needs Outcome Date'
                else:
                    return ""
            elif Level in leveltypes and Over62 == "Yes" and Type == "NYCHA Housing Termination" and Disposition == "Closed":
                if Outcome == '' and OutcomeDate == '':
                    return 'Needs Outcome & Date'
                elif  Outcome == '':
                    return 'Needs Outcome'
                elif OutcomeDate == '':
                    return 'Needs Outcome Date'
                else:
                    return ""
            else:
                return ''
        
        df['Outcome Tester'] = df.apply(lambda x: OutcomeTesterClean(x['Over 62?'],x['Housing Outcome'],x['Housing Outcome Date'],x['Housing Level of Service'],x['Housing Type Of Case'],x['Case Disposition']), axis = 1)
        

        #Test if Poverty Percentage > 1000%
        
        def PovertyPercentTester (PovertyPercent,WaiverCategory):
            if PovertyPercent > 1000 and WaiverCategory != "Income Waiver":
                return "Needs Income Review"
            else:
                return ""
        df['Poverty Percent Tester'] = df.apply(lambda x: PovertyPercentTester(x['Percentage of Poverty'],x['Housing TRC HRA Waiver Categories']), axis = 1)
        
        #Test # of adults (can't be 0)
        
        def AdultTester (HouseholdAdults,BirthDate):
            if BirthDate == '':
                return 'Needs Date of Birth'
            
            elif BirthDate != '':
                BirthDateMonth = int(BirthDate[:2])
                BirthDateDay = int(BirthDate[3:5])
                BirthDateYear = int(BirthDate[6:])
                
                Today = datetime.today()
                Today = Today.strftime("%m/%d/%Y")
                
                TodayMonth = int(Today[:2])
                TodayDay = int(Today[3:5])
                TodayYear = int(Today [6:])
                
                if TodayMonth == 1:
                    TodayMonth = 13
                
                #BirthDateConstruct = int(BirthDateYear + BirthDateMonth + BirthDateDay)
                #TodayConstruct = int(TodayYear + TodayMonth + TodayDay)
                
                if TodayYear - BirthDateYear > 18:
                    if HouseholdAdults == 0:
                        return "Needs Over-18 Household Member"
                    else:
                        return ""
                elif TodayYear - BirthDateYear == 18 and TodayMonth - 1 >= BirthDateMonth:
                    if HouseholdAdults == 0:
                        return "Needs Over-18 Household Member"
                    else:
                        return ""
                else: 
                    return "Client Needs to be Over 18"
            else:
                return ""
        df['Over-18 Tester'] = df.apply(lambda x: AdultTester(x['Number of People 18 and Over'],x['Date of Birth']), axis = 1)
        
        #date of waiver approval & waiver categories - if there's something in one but not the other, then flag it. 
        
        def WaiverTester (WaiverType,WaiverDate):
            if WaiverType != "" and WaiverDate == "":
                return "Needs Waiver Date"
            elif WaiverDate != "" and WaiverType == "":
                return "Needs Waiver Type"
            else:
                return ""
        df['Waiver Tester'] = df.apply(lambda x: WaiverTester(x['Housing TRC HRA Waiver Categories'],x['Housing Date Of Waiver Approval']), axis = 1)
        
        def FundingTester (PrimaryFunding,SecondaryFunding):
            if PrimaryFunding == SecondaryFunding:
                return ""
            elif "3011" in SecondaryFunding or "3018" in SecondaryFunding or "3111" in SecondaryFunding or "3112" in SecondaryFunding or "3113" in SecondaryFunding or "3114" in SecondaryFunding or "3115" in SecondaryFunding or "3121" in SecondaryFunding or "3122" in SecondaryFunding or "3123" in SecondaryFunding or "3124" in SecondaryFunding or "3125" in SecondaryFunding:
                return "Needs Funding Code Review"
            else:
                return ""
            
                
        df['Funding Tester'] = df.apply(lambda x: FundingTester(x['Primary Funding Code'],x['Secondary Funding Codes']), axis = 1)
        
        
        #COVID Modifications - make the testers blank if it's an advice only pre-3/1 case!
        """
        #Differentiate pre- and post- 3/1/20 eligibility date cases
           
        df['EligConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        df['OpenedConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']), axis=1)
        
        def PreThreeOne(EligibilityDate,OpenedDate):
            if isinstance(EligibilityDate, int) == False:
                return 'Eligbility Date Missing'
            elif EligibilityDate < 20200301:
                return "Yes"
            elif EligibilityDate >= 20200301:
                return "No"
        df['Pre-3/1/20 Elig Date?'] = df.apply(lambda x: PreThreeOne(x['EligConstruct'],x['OpenedConstruct']), axis=1)
        
        
        def NeedsRedactingTester(LevelOfService,PreThreeOne,FundingCodeSorter,LegalProblemCode,EligDate,DateOpened):
            if isinstance(EligDate, int) == False:
                EligDate = 0            
                
                if DateOpened >= 20200301 and LevelOfService.startswith("Advice") == True:
                    return "Needs Redacting"
                elif DateOpened >= 20200301 and LevelOfService.startswith("Brief") == True and FundingCodeSorter == "UAHPLP":
                    return "Needs Redacting"
                elif LegalProblemCode.startswith("6") == False and DateOpened <= 20200930 and DateOpened >= 20200301:
                    return "Needs Redacting"
                else:
                    return ""
            elif LevelOfService.startswith("Advice") == True and PreThreeOne == "No":
                return "Needs Redacting"
            elif LevelOfService.startswith("Brief") == True and PreThreeOne == "No" and FundingCodeSorter == "UAHPLP":
                return "Needs Redacting"
            elif LegalProblemCode.startswith("6") == False and EligDate <= 20200930 and EligDate >= 20200301:
                return "Needs Redacting"
            else:
                return ""
        df['Post-3/1 Limited Service Tester'] = df.apply(lambda x: NeedsRedactingTester(x['Housing Level of Service'],x['Pre-3/1/20 Elig Date?'],x['Funding Code Sorter'],x['Legal Problem Code'],x['EligConstruct'],x['OpenedConstruct']), axis=1)

        #CovidException testers to erase clean-up requests

        def AllHousingRedactForCovid(LimitedServiceTester, ToRedact):
            if LimitedServiceTester == "Needs Redacting":
                return ""
            else:   
                return ToRedact
        
        df['PA # Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['PA # Tester']), axis=1)
        
        df['SS # Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['SS # Tester']), axis=1)
        
        df['Case Number Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Case Number Tester']), axis=1)
        
        df['Rent Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Rent Tester']), axis=1)       
       
        
        df['Years in Apartment Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Years in Apartment Tester']), axis=1)
        
        df['Referral Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Referral Tester']), axis=1)

        
        df['Posture Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Posture Tester']), axis=1)
        
        df['Unit Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Unit Tester']), axis=1)
        
        df['Regulation Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Regulation Tester']), axis=1)
        
        df['Subsidy Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Subsidy Tester']), axis=1)
        
        df['Outcome Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Outcome Tester']), axis=1)
        
        df['Housing Services Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Housing Services Tester']), axis=1)
        
        df['Housing Activity Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Housing Activity Tester']), axis=1)
        
        df['Release & Elig Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Release & Elig Tester']), axis=1)
        
        df['Building Case Tester'] = df.apply(lambda x: AllHousingRedactForCovid(x['Post-3/1 Limited Service Tester'], x['Building Case Tester']), axis=1)
        """
        #Tester Tester
        
        def TesterTester(ReleaseEligTester,PostureTester,CaseNumberTester,HousingTypeTester,HousingLevelTester,BuildingCaseTester,ReferralTester,PATester,SSTester,UnitTester,RegulationTester,SubsidyTester,YearsAptTester,LanguageTester,HousingActivityTester,HousingServicesTester,OutcomeTester,OverEighteenTester,PovertyPercentTester,WaiverTester,FundingTester):
            if "Needs" in ReleaseEligTester:
                return "Needs Cleanup"
            elif "Needs" in PostureTester:
                return "Needs Cleanup"
            elif "Needs" in CaseNumberTester:
                return "Needs Cleanup"
            elif "Needs" in HousingTypeTester:
                return "Needs Cleanup"
            elif "Needs" in HousingLevelTester:
                return "Needs Cleanup"
            elif "Needs" in BuildingCaseTester:
                return "Needs Cleanup"
            elif "Needs" in ReferralTester:
                return "Needs Cleanup"
            elif "Needs" in PATester:
                return "Needs Cleanup"
            elif "Needs" in SSTester:
                return "Needs Cleanup"
            elif "Needs" in UnitTester:
                return "Needs Cleanup"
            elif "Needs" in RegulationTester:
                return "Needs Cleanup"
            elif "Needs" in SubsidyTester:
                return "Needs Cleanup"
            elif "Needs" in YearsAptTester:
                return "Needs Cleanup"
            elif "Needs" in LanguageTester:
                return "Needs Cleanup"
            elif "Needs" in HousingActivityTester:
                return "Needs Cleanup"
            elif "Needs" in HousingServicesTester:
                return "Needs Cleanup"
            elif "Needs" in OutcomeTester:
                return "Needs Cleanup"
            elif "Needs" in OverEighteenTester:
                return "Needs Cleanup"
            elif "Needs" in PovertyPercentTester:
                return "Needs Cleanup"
            elif "Needs" in WaiverTester:
                return "Needs Cleanup"
            elif "Needs" in FundingTester:
                return "Needs Cleanup"
            else:
                return "No Cleanup Necessary"
        
        
        df['Tester Tester'] = df.apply(lambda x: TesterTester(x['Release & Elig Tester'],x['Posture Tester'],x['Case Number Tester'],x['Housing Type Tester'],x['Housing Level Tester'],x['Building Case Tester'],x['Referral Tester'],x['PA # Tester'],x['SS # Tester'],x['Unit Tester'],x['Regulation Tester'],x['Subsidy Tester'],x['Years in Apartment Tester'],x['Language Tester'],x['Housing Activity Tester'],x['Housing Services Tester'],x['Outcome Tester'],x['Over-18 Tester'],x['Poverty Percent Tester'],x['Waiver Tester'],x['Funding Tester']), axis = 1)
       
        
        #sort by case handler
        
        df = df.sort_values(by=['Primary Advocate'])
        df = df.sort_values(by=['Assigned Branch/CC'])
        df = df.sort_values(by=['Tester Tester'])
        
        
        #Put everything in the right order
        
        df = df[['Hyperlinked CaseID#',
        "Assigned Branch/CC",
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
        'Release & Elig Tester',
        "Housing Income Verification",
        #'Income Verification Tester',
        "Housing Posture of Case on Eligibility Date",'Posture Tester',
        "Gen Case Index Number",'Case Number Tester',  
        "Housing Type Of Case",'Housing Type Tester',
        "Housing Level of Service",'Housing Level Tester',
        "Close Reason",
        "Housing Building Case?",'Building Case Tester',
        "Primary Funding Code",
        "Secondary Funding Codes",
        "Funding Tester",
        "Housing Total Monthly Rent",
        #'Rent Tester',
        "Referral Source",'Referral Tester',
        "Gen Pub Assist Case Number",'PA # Tester',
        "Social Security #","SS # Tester",
        "Housing Number Of Units In Building",'Unit Tester',
        "Housing Form Of Regulation",'Regulation Tester',
        "Housing Subsidy Type",'Subsidy Tester',
        "Housing Years Living In Apartment",'Years in Apartment Tester',
        "Language",'Language Tester',
        "Housing Activity Indicators",'Housing Activity Tester',
        "Housing Services Rendered to Client",'Housing Services Tester',
        "Housing Outcome",'Outcome Tester',"Housing Outcome Date",
        "Assigned Branch/CC",
        "Number of People 18 and Over","Number of People under 18","Over-18 Tester",
        "Date of Birth",
        "Over 62?",
        "Percentage of Poverty","Poverty Percent Tester",
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
        #'Post-3/1 Limited Service Tester'
        

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
                medium_problem_format = workbook.add_format({'bg_color':'cyan'})
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
                ws.set_column('C:C',32)
                ws.autofilter('B1:ZZ1')
                ws.freeze_panes(1, 2)
                ws.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'No Release - Remove Elig Date',
                                                 'format': bad_problem_format})
                ws.conditional_format('AM2:AM100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Valid',
                                                 'format': bad_problem_format})
                ws.conditional_format('C2:BO100000',{'type': 'text',
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
