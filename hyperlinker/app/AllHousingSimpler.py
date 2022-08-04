from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory 
from app import app, db, DataWizardTools, HousingToolBox 
from app.models import User, Post 
from app.forms import PostForm 
from werkzeug.urls import url_parse 
from datetime import datetime 
import pandas as pd 
 
 
@app.route("/AllHousingSimpler", methods=['GET', 'POST']) 
def AllHousingSimpler(): 
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
         
        #Housing Type of Case Eviction-Types: 
        evictiontypes = ['Holdover','Non-payment','Illegal Lockout','NYCHA Housing Termination'] 
        #Highest Level of Service Reps 
        leveltypes = ['Representation - State Court','Representation - Federal Court'] 
         
        df['Agency'] = "LSNYC" 
 
        df['Assigned Branch/CC'] = df.apply(lambda x : DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis=1)    
         
        #Puts dates in a format that can be compared (yyyymmdd) 
        df['EDate Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']),axis=1) 

         
        def EDateConstructMaker (EDC): 
            if EDC == "": 
                return 0 
            else: 
                return EDC 
                 
        df['EDate Construct'] = df.apply(lambda x: EDateConstructMaker(x['EDate Construct']), axis=1) 
        df['Today'] = datetime.now() 
        df['TodayConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Today']), axis=1) 
         
        def ReleaseAndEligTester(HRARelease,EligibilityDate, EliDC, TdC, LoS, Close, PrimaryFunding): 
            UAHPLPFund = ['3121 Universal Access to Counsel – (UAC)','3122 Universal Access to Counsel – (UAC)','3123 Universal Access to Counsel – (UAC)','3124 Universal Access to Counsel – (UAC)','3125 Universal Access to Counsel – (UAC)','3111 HPLP-Homelessness Prevention Law Project','3112 HPLP-Homelessness Prevention Law Project','3113 HPLP-Homelessness Prevention Law Project','3114 HRA-HPLP-Homelessness Prevention Law Project','3115 HPLP-Homelessness Prevention Law Project'] 
            LoSEmpty = ["", "Hold For Review"] 
            if 20220701 > EliDC > 20000101: 
                return "Liz to check or remove old EDate Case" 
            elif 20000101 >= EliDC > 0: 
                return "Liz to review old EDate" 
            elif TdC < EliDC: 
                return "Liz to review future EDate" 
            elif HRARelease == 'Yes' and EligibilityDate != '': 
                return 'Reportable' 
            elif HRARelease == 'Yes' and EligibilityDate == '': 
                return "Needs Eligibility Date" 
            elif HRARelease != "Yes": 
                if EligibilityDate != '': 
                    return "Needs HRA Release" 
                else:    
                    return 'Needs Elig & Release' 
            else: 
                return 'Unknown Answer Category' 
        df['Unreportable'] = df.apply(lambda x: ReleaseAndEligTester(x['HRA Release?'],x['HAL Eligibility Date'],x['EDate Construct'],x['TodayConstruct'],x['Housing Level of Service'],x['Close Reason'],x['Primary Funding Code']), axis=1) 
         
        #Has to have a Housing Type of Case **Tester column removed 
        def HousingTypeClean (HousingType): 
            if HousingType == '': 
                return 'Needs Housing Type of Case' 
            else: 
                return HousingType 
         
        df['Housing Type Of Case'] = df.apply(lambda x: HousingTypeClean(x['Housing Type Of Case']), axis=1) 
         
        #Has to have a Housing Level of Service **Tester column removed 
         
        def HousingLevelClean (HousingLevel): 
            if HousingLevel == '': 
                return 'Needs Level of Service'
            elif HousingLevel == 'Hold For Review':
                return 'Needs Updated Level of Service'
            else: 
                return HousingLevel 
         
        df['Housing Level of Service'] = df.apply(lambda x: HousingLevelClean(x['Housing Level of Service']), axis=1) 
        
        #Has to say whether or not it's a building case **Tester column removed 
        def BuildingCaseClean (BuildingCase): 
            if BuildingCase == '': 
                return 'Needs Building Case Answer' 
            if BuildingCase == 'Prefer Not To Answer': 
                return 'Needs Building Case Answer' 
            else: 
                return BuildingCase 
       
        df['Housing Building Case?'] = df.apply(lambda x: BuildingCaseClean(x['Housing Building Case?']), axis=1) 
         
        #Referral Source Can't Be Blank **Tester column removed 
        def ReferralClean (Referral,FundingSource): 
            if Referral == '': 
                return 'Needs Referral Source' 
            elif FundingSource == '3011 TRC FJC Initiative' and Referral != 'FJC Housing Intake': 
                return "Says '" + Referral + "'," + ' Needs to be FJC if funded 3011'
            elif FundingSource != '3011 TRC FJC Initiative' and Referral == 'FJC Housing Intake':
                return "Says '" + Referral + "'," + ' Needs funding code 3011 if FJC Referred'
            else: 
                return Referral 
         
        df['Referral Source'] = df.apply(lambda x: ReferralClean(x['Referral Source'],x['Primary Funding Code']), axis=1) 
         
        #number of units in building can't be 0 or written with letters **Tester column removed, decimals eliminated 
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
                return float(Units) 
        df['Housing Number Of Units In Building'] = df.apply(lambda x: UnitsClean(x['Housing Number Of Units In Building']), axis=1) 
         
        #Housing form of regulation can't be blank **Tester column removed 
        def RegulationClean (Regulation): 
            if Regulation == '': 
                return 'Needs Form of Regulation' 
            else: 
                return Regulation 
        df['Housing Form Of Regulation'] = df.apply(lambda x: RegulationClean(x['Housing Form Of Regulation']), axis=1) 
         
        #Housing subsidy can't be blank (can be none) **Tester column removed 
        def SubsidyClean (Subsidy): 
            if Subsidy == '': 
                return 'Needs Type of Subsidy' 
            else: 
                return Subsidy 
        df['Housing Subsidy Type'] = df.apply(lambda x: SubsidyClean(x['Housing Subsidy Type']), axis=1) 
        
        df['BDay Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date of Birth']),axis=1)        
        
        #Years in Apartment Can't be 0 (can be -1) **in process, close age problem 
        def YearsClean (Years,Age,DOB,TDC,BDC): 
            
            if isinstance(Years, int) == False:
                return str(Years) + ", Answer needs to be integer"
            elif DOB == "":
                return str(Years) + ", Needs client Date of Birth"
            elif TDC - BDC < 180000:
                return str(Years) + ", Needs adult client DOB"
            elif Years*10000 > 20230630-BDC:
                return str(int(Years)) + ", Needs Years not exceeding client age"
            elif Years == 0: 
                return '0, Needs Years In Apartment' 
            elif Years > 99: 
                return str(int(Years)) + ', Needs Valid Number' 
            elif Years < -1: 
                return str(int(Years)) + ', Unless <1 yr in apt, Needs positive number'
            else: 
                return Years
        df['Housing Years Living In Apartment'] = df.apply(lambda x: YearsClean(x['Housing Years Living In Apartment'],x['Age at Intake'],x['Date of Birth'],x['TodayConstruct'],x['BDay Construct']), axis=1) 
         
        #Years > Age+1: return str(int(Years)) + ', Needs years not exceeding client age'
        #Language Can't be Blank or Unknown **in process, Not entered? 
        def LanguageClean (Language): 
            if Language == '' or Language == 'Unknown': 
                return 'Needs Language' 
            else: 
                return Language 
        df['Language'] = df.apply(lambda x: LanguageClean(x['Language']), axis=1) 
         
        #Housing Income Verification can't be blank or none and other stuff with kids and poverty level and you just give up if it's closed 
         
        #df['Income Verification Tester'] = df.apply(lambda x: HousingToolBox.IncomeVerificationClean(x['Housing Income Verification'], x['Number of People under 18'], x['Percentage of Poverty'],x['Case Disposition']), axis=1) 
         
        #Test if social security number is correct format (or ignore it if there's a valid PA number) 
        # if First3 == '000' and Middle2 == '00' and Last4 == '0000': return CaseNum + ' Needs SS#' 
        def SSNumClean (CaseNum,PACit): 
            CaseNum = str(CaseNum) 
            First3 = CaseNum[0:3] 
            Middle2 = CaseNum[4:6] 
            Last4 = CaseNum[7:11] 
            FirstDash = CaseNum[3:4] 
            SecondDash = CaseNum[6:7] 
            
            if str.isnumeric(First3) == True and str.isnumeric(Middle2) == True and str.isnumeric(Last4) == True and FirstDash == '-' and SecondDash == '-':  
                return CaseNum 
            elif "citizen" in str(PACit): 
                return "Has No SSN" 
            else: 
                return "Needs SS#" 
                 
        df['Social Security #'] = df.apply(lambda x: SSNumClean(x['Social Security #'],x['Gen Pub Assist Case Number']), axis=1) 
         
         
        #PA Tester ** needs updating, can possibly be done
                
        def PATesterClean (PANumber,SSTester): 
                         
            PANumber = str(PANumber) 
            LastCharacter = PANumber[-1:] 
            PenultimateCharater = PANumber[-2:-1] 
            SecondCharacter = PANumber [1:2] 
            if SSTester == '': 
                return 'Unnecessary due to SS#' 
            elif PANumber == '' or PANumber == 'None' or PANumber == 'NONE' or SecondCharacter == 'o': 
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
 
         
        df['PA # Tester'] = df.apply(lambda x: PATesterClean(x['Gen Pub Assist Case Number'],x['Social Security #']), axis=1)
         
         
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
             
            if Level == 'Advice' or Level == 'Brief Service' or Level == 'Out-of-Court Advocacy' or Level == 'Hold For Review' or Level.startswith('UAC') == True: 
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
                 
        def ServicesTesterClean(HousingServices,Level,Type): 
             
            if Level == 'Representation - State Court' or Level == 'Representation - Admin. Agency' or Level == 'Representation - Federal Court' or Level == 'UAC Out of Court Advocacy WITH Retainer': 
                if HousingServices == '': 
                    return 'Needs Services Rendered' 
                else: 
                    return HousingServices 
            else: 
                return HousingServices 
 
        df['Housing Services Rendered to Client'] = df.apply(lambda x: ServicesTesterClean(x['Housing Services Rendered to Client'],x['Housing Level of Service'],x['Housing Type Of Case']), axis = 1) 
         
         
        #Test Housing Activity Indicator - can't be blank for cases that are full rep state or full rep federal(housing level of service) and eviction cases(housing type of case: non-payment holdover illegal lockout or nycha housing termination) **Tester column removed 
         
        def ActivityTesterClean(HousingActivity,Level,Type): 
            if Level == 'Representation - State Court' or Level == 'Representation - Admin. Agency' or Level == 'Representation - Federal Court': 
                if Type in evictiontypes and HousingActivity == '': 
                    return 'Needs Activity Indicator' 
                else: 
                    return HousingActivity   
            else: 
                return HousingActivity 
        df['Housing Activity Indicators'] = df.apply(lambda x: ActivityTesterClean(x['Housing Activity Indicators'],x['Housing Level of Service'],x['Housing Type Of Case']), axis = 1) 
         
        #Housing Posture of Case can't be blank if case is Full Rep Eviction **Tester column removed 
         
        def PostureClean (Posture,Type,Level): 
            if Level == 'Representation - State Court' or Level == 'Representation - Admin. Agency' or Level == 'Representation - Federal Court': 
                if Type in evictiontypes and Posture == '': 
                    return 'Needs Posture of Case' 
                else: 
                    return Posture  
            else: 
                return Posture 
         
        df['Housing Posture of Case on Eligibility Date'] = df.apply(lambda x: PostureClean(x['Housing Posture of Case on Eligibility Date'],x['Housing Type Of Case'],x['Housing Level of Service']), axis=1) 
         
         
        #Outcome Tester - needs outcome and date for eviction cases that are full rep at state or federal level (not admin)**complete, Tester split in 2 
         
        def OutcomeClean (Outcome,Level,Type,Disposition): 
             
            if Level == 'Representation - State Court' or Level == 'Representation - Admin. Agency' or Level == 'Representation - Federal Court' or Level == 'UAC Out of Court Advocacy WITH Retainer': 
                if Type in evictiontypes and Disposition == "Closed": 
                    if  Outcome == '': 
                        return 'Needs Outcome' 
                    else: 
                        return Outcome 
                else:  
                    return Outcome 
            else: 
                return Outcome 
         
        df['Housing Outcome'] = df.apply(lambda x: OutcomeClean(x['Housing Outcome'],x['Housing Level of Service'],x['Housing Type Of Case'],x['Case Disposition']), axis = 1) 
         
        def OutcomeDateClean (OutcomeDate,Level,Type,Disposition): 
             
            if Level == 'Representation - State Court' or Level == 'Representation - Admin. Agency' or Level == 'Representation - Federal Court' or Level == 'UAC Out of Court Advocacy WITH Retainer': 
                if Type in evictiontypes and Disposition == "Closed": 
                    if OutcomeDate == '': 
                        return 'Needs Outcome Date' 
                    else: 
                        return OutcomeDate 
                else:  
                    return OutcomeDate 
            else: 
                return OutcomeDate 
         
        df['Housing Outcome Date'] = df.apply(lambda x: OutcomeDateClean(x['Housing Outcome Date'],x['Housing Level of Service'],x['Housing Type Of Case'],x['Case Disposition']), axis = 1) 
         
         
        #Test # of adults (can't be 0)

        def AdultTester (HouseholdAdults,Under18s,BirthDate): 
            if HouseholdAdults == 0 and Under18s == 0: 
                    return "Needs Household Members" 
             
            elif BirthDate == '': 
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
        df['Over-18 Tester'] = df.apply(lambda x: AdultTester(x['Number of People 18 and Over'],x['Number of People under 18'],x['Date of Birth']), axis = 1) 
        
        def DOBTester (BirthDate, TDC, BDC):
            if BirthDate == "":
                return "Needs Client DOB"
            elif TDC - BDC < 180000:
                return str(BirthDate) + ", Client needs to be 18+"
            else:
                return BirthDate
                
        df['Date of Birth'] = df.apply(lambda x: DOBTester(x['Date of Birth'],x['TodayConstruct'],x['BDay Construct']), axis = 1)
        
        def Over18Tester (Over, Under):
            if Over == 0 and Under == 0:
                return str(int(Over)) + ", Needs Household Members"
            elif Over == 0 and Under != 0:
                return str(int(Over)) + ", Needs Over 18 household member"
            elif Over + Under >= 18:
                return str(int(Over)) + ", Possible error needs attention, adults + children > 18"
            else:
                return Over
                
        df['Number of People 18 and Over'] = df.apply(lambda x: Over18Tester(x['Number of People 18 and Over'],x['Number of People under 18']), axis = 1)       
         
        #date of waiver approval & waiver categories - if there's something in one but not the other, then flag it. **in process, zip code? 
         
        def WaiverTester (WaiverType,WaiverDate,Poverty,RefSou,Type,LoS,PrimaryFunding,EDC): 
            UAHPLPFund = ['3121 Universal Access to Counsel – (UAC)','3122 Universal Access to Counsel – (UAC)','3123 Universal Access to Counsel – (UAC)','3124 Universal Access to Counsel – (UAC)','3125 Universal Access to Counsel – (UAC)','3111 HPLP-Homelessness Prevention Law Project','3112 HPLP-Homelessness Prevention Law Project','3113 HPLP-Homelessness Prevention Law Project','3114 HRA-HPLP-Homelessness Prevention Law Project','3115 HPLP-Homelessness Prevention Law Project']
            global HousingFundingCodes
            HousingFundingCodes= ['3121 Universal Access to Counsel – (UAC)','3122 Universal Access to Counsel – (UAC)','3123 Universal Access to Counsel – (UAC)','3124 Universal Access to Counsel – (UAC)','3125 Universal Access to Counsel – (UAC)','3111 HPLP-Homelessness Prevention Law Project','3112 HPLP-Homelessness Prevention Law Project','3113 HPLP-Homelessness Prevention Law Project','3114 HRA-HPLP-Homelessness Prevention Law Project','3115 HPLP-Homelessness Prevention Law Project','3018 Tenant Rights Coalition (TRC)','3011 TRC FJC Initiative'] 
            if Poverty < 201: 
                if WaiverType == "" and WaiverDate == "":
                    return "Income Eligible, <201%"
                else:
                    return "Waiver fields can be cleared, income eligible case does not need waiver"
            elif PrimaryFunding == "3011 TRC FJC Initiative":
                if WaiverDate == "" and WaiverType == "":
                    return "Needs waiver type, 'FJC Waiver' and waiver date '11/28/2016'"
                elif WaiverDate != "11/28/2016": 
                    return "Needs waiver date '11/28/2016'"
                elif WaiverType != "FJC Waiver":
                    return "Needs waiver type 'FJC Waver' if funded 3011"
                else:
                    return "FJC waiver, fine"
            elif PrimaryFunding != "3011 TRC FJC Initative" and WaiverType == "FJC Waiver":
                return "Needs change, only cases referred by FJC (3011) can have FJC waiver"
            elif WaiverType != "" and WaiverDate != "": 
                return "Already has waiver"
            elif WaiverDate != "" and WaiverType == "": 
                return "Needs Waiver Type" 
            elif WaiverType != "" and WaiverDate == "":
                return "Needs Waiver Date"
            elif PrimaryFunding not in HousingFundingCodes:
                return "Not housing case, waiver not needed at this time"
            elif "Needs" in LoS:
                return "Missing Level of Service, waiver need unknown"
            elif "Needs" in Type:
                return "Missing Type of Case, waiver need unknown"
            elif PrimaryFunding in HousingFundingCodes:
                if LoS == "Advice":
                    return "Housing Advice, no waiver needed at this time"
                elif PrimaryFunding in UAHPLPFund and LoS == "Brief Service":
                    return "UAHPLP Brief, no waiver needed at this time"
                elif Type == "Illegal Lockout":
                    return "Needs update, ILO has categorical income waiver with date '09/20/2021'"
                else:
                    return "Over income housing case, may need waiver requested"
            else: 
                return "Possible waiver request needed - Liz to check" 
        df['Waiver Tester'] = df.apply(lambda x: WaiverTester(x['Housing TRC HRA Waiver Categories'],x['Housing Date Of Waiver Approval'],x['Percentage of Poverty'],x['Referral Source'],x['Housing Type Of Case'],x['Housing Level of Service'],x['Primary Funding Code'],x['EDate Construct']), axis = 1) 
         
        #Test if Poverty Percentage > 1000% **Tester column removed 
         
        def PovertyPercentTester (PovertyPercent,WaiverCategory): 
            if WaiverCategory == "Income Waiver" or WaiverCategory == "FJC Waiver": 
                return PovertyPercent 
            elif PovertyPercent > 1000 : 
                return str(PovertyPercent) + " May need Income Review" 
            else: 
                return PovertyPercent 
        df['Percentage of Poverty'] = df.apply(lambda x: PovertyPercentTester(x['Percentage of Poverty'],x['Housing TRC HRA Waiver Categories']), axis = 1) 
         
        #**in process, FJC?, secondary funding codes answer checked 
        #source of unique 'set' below - https://stackoverflow.com/questions/12897374/get-unique-values-from-a-list-in-python 
        def FundingTester (PrimaryFunding,SecondaryFunding): 
            GFCodes = ['2157 OCA-City-wide Civil Legal Services Grant','3020 CLS-Civil Legal Services','4000 LSC - Basic Grant','4100 IOLA - General','5221 SSUSA-Single Stop USA'] 
            GFCodesAll = GFCodes + ['5510 CB9 Manhattanville-West Harlem Tenant Advocacy Project'] 
            ManhattanFundingCodes= ['3123 Universal Access to Counsel – (UAC)','3115 HPLP-Homelessness Prevention Law Project'] 
            #WholeUniCodes= HousingFundingCodes + GFCodesAll 
            SplitSecondaryList = list(set(SecondaryFunding.split(', '))) 
            #AllFundingList = list(SplitSecondaryList+PrimaryFunding)
            #SplitSecondarySet = set(SplitSecondaryList) 
            #SplitSecondaryList = list(SplitSecondarySet) 
            #print(SplitSecondaryList) 
            #print(GFCodesAll) 
             
            if SecondaryFunding != "": 
                if len(SplitSecondaryList) > 1: 
                    if PrimaryFunding in HousingFundingCodes:
                        if PrimaryFunding in SplitSecondaryList:
                            if SplitSecondaryList[0] in GFCodesAll or SplitSecondaryList[1] in GFCodesAll:
                                return "Regular housing double, likely okay"
                            else :
                                return "Needs review. Too many funding codes, housing problem"
                        else:
                            return "Needs review. Too many funding codes, 3 different codes"
                    else:
                        return "Needs review. Too many funding codes, 3 diff, not prim. housing"
                elif PrimaryFunding in GFCodesAll and SplitSecondaryList[0] in GFCodesAll: 
                    return "All General, Fine" 
                elif PrimaryFunding not in HousingFundingCodes: 
                    if SplitSecondaryList[0] in HousingFundingCodes: 
                        return "Needs review, HRA Funding code can only be secondary to self" 
                    else: 
                        return "Targeted" 
                elif PrimaryFunding in ManhattanFundingCodes and "5510 CB9 Manhattanville-West Harlem Tenant Advocacy Project" in SplitSecondaryList: 
                    return "Manhattan Fine" 
                elif PrimaryFunding in HousingFundingCodes and SplitSecondaryList[0] in GFCodes: 
                    return "General 2nd Fine" 
                elif PrimaryFunding in HousingFundingCodes and PrimaryFunding != SplitSecondaryList[0] and SplitSecondaryList[0] in HousingFundingCodes: 
                    return "Needs Funding Code Review" 
                elif PrimaryFunding not in SecondaryFunding and SecondaryFunding not in GFCodes: 
                    return "Needs Secondary Funding Code Review" 
                elif PrimaryFunding == SplitSecondaryList[0]: 
                    return "Same code, Fine" 
                else: 
                    return "Something is Wrong" 
            else: 
                return "Primary only, Fine" 
            '''elif SecondaryFunding not in HousingFundingCodes and SecondaryFunding not in GFCodes: 
                return "Needs Funding Code Review" 
            else: 
                return ""''' 
             
                 
        df['Funding Tester'] = df.apply(lambda x: FundingTester(x['Primary Funding Code'],x['Secondary Funding Codes']), axis = 1) 
         
        def HousingTabAssigner (PrimaryFunding,SecondaryFunding):
            UAHPLP = ['3121 Universal Access to Counsel – (UAC)','3122 Universal Access to Counsel – (UAC)','3123 Universal Access to Counsel – (UAC)','3124 Universal Access to Counsel – (UAC)','3125 Universal Access to Counsel – (UAC)','3111 HPLP-Homelessness Prevention Law Project','3112 HPLP-Homelessness Prevention Law Project','3113 HPLP-Homelessness Prevention Law Project','3114 HRA-HPLP-Homelessness Prevention Law Project','3115 HPLP-Homelessness Prevention Law Project']
            TRC = ['3018 Tenant Rights Coalition (TRC)','3011 TRC FJC Initiative'] 
            Housing = UAHPLP + TRC
            SplitSecondaryList = list(set(SecondaryFunding.split(', '))) 
            #print (SplitSecondaryList)
            if PrimaryFunding in UAHPLP:
                return "UA HPLP"
            elif PrimaryFunding in TRC:
                return "TRC"
            elif SecondaryFunding == "":
                return "Other Cases"
            elif len(SplitSecondaryList) == 1:
                if SplitSecondaryList[0] in UAHPLP:
                    return "UA HPLP"
                elif SplitSecondaryList[0] in TRC:
                    return "TRC"
                else:
                    return "Other Cases"
            elif SplitSecondaryList[0] in UAHPLP or SplitSecondaryList[1] in UAHPLP:
                return "UA HPLP"
            elif SplitSecondaryList[0] in TRC or SplitSecondaryList[1] in TRC:
                return "TRC"
            else:
                return "Other Cases"
        
        df['Housing Tab Assignment'] = df.apply(lambda x: HousingTabAssigner(x['Primary Funding Code'],x['Secondary Funding Codes']), axis = 1)
       
         
        
        #Tester Tester 
         
        def TesterTester(ReleaseEligTester,PostureTester,CaseNumberTester,HousingTypeTester,HousingLevelTester,BuildingCaseTester,ReferralTester,PATester,SSTester,UnitTester,RegulationTester,SubsidyTester,YearsAptTester,LanguageTester,HousingActivityTester,HousingServicesTester,OutcomeTester,OutcomeDateTester,OverEighteenTester,PovertyPercentTester,WaiverTester,FundingTester): 
             
            if "Needs" in PostureTester: 
                return "Needs Cleanup" 
            elif "Needs" in HousingTypeTester: 
                return "Needs Cleanup" 
            elif "Needs" in HousingLevelTester: 
                return "Needs Cleanup" 
            elif "Needs" in BuildingCaseTester: 
                return "Needs Cleanup" 
            elif "Needs" in ReferralTester: 
                return "Needs Cleanup" 
            elif "Needs" in SSTester: 
                return "Needs Cleanup" 
            elif "Needs" in str(UnitTester): 
                return "Needs Cleanup" 
            elif "Needs" in RegulationTester: 
                return "Needs Cleanup" 
            elif "Needs" in SubsidyTester: 
                return "Needs Cleanup" 
            elif "Needs" in str(YearsAptTester): 
                return "Needs Cleanup" 
            elif "Needs" in LanguageTester: 
                return "Needs Cleanup" 
            elif "Needs" in HousingActivityTester: 
                return "Needs Cleanup" 
            elif "Needs" in HousingServicesTester: 
                return "Needs Cleanup" 
            elif "Needs" in OutcomeTester: 
                return "Needs Cleanup" 
            elif "Needs" in OutcomeDateTester: 
                return "Needs Cleanup" 
            elif "Needs" in OverEighteenTester: 
                return "Needs Cleanup" 
            elif "Needs" in str(PovertyPercentTester): 
                return "Needs Cleanup" 
            elif "Needs" in str(WaiverTester): 
                return "Needs Cleanup" 
            elif "Needs" in FundingTester or "secondary" in FundingTester or "many" in FundingTester or "Targeted" in FundingTester: 
                return "Needs Cleanup" 
            if "Needs" in str(ReleaseEligTester): 
                return "Needs Cleanup" 
            elif "Needs" in CaseNumberTester: 
                return "Untested" 
            elif "Needs" in PATester: 
                return "Untested" 
            else: 
                return "No Cleanup Necessary" 
         
         
        df['Tester Tester'] = df.apply(lambda x: TesterTester(x['Unreportable'],x['Housing Posture of Case on Eligibility Date'],x['Case Number Tester'],x['Housing Type Of Case'],x['Housing Level of Service'],x['Housing Building Case?'],x['Referral Source'],x['PA # Tester'],x['Social Security #'],x['Housing Number Of Units In Building'],x['Housing Form Of Regulation'],x['Housing Subsidy Type'],x['Housing Years Living In Apartment'],x['Language'],x['Housing Activity Indicators'],x['Housing Services Rendered to Client'],x['Housing Outcome'],x['Housing Outcome Date'],x['Over-18 Tester'],x['Percentage of Poverty'],x['Waiver Tester'],x['Funding Tester']), axis = 1) 
        
         
        #sort by case handler 
         
        df = df.sort_values(by=['Primary Advocate']) 
        df = df.sort_values(by=['Assigned Branch/CC']) 
        df = df.sort_values(by=['Tester Tester']) 
         
         
        #Put everything in the right order **link the second step to this list**
         
        df = df[['Hyperlinked CaseID#', 
        "Assigned Branch/CC", 
        #'Tester Tester', 
        "Primary Advocate", 
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
        #"Pre-12/1/21 Elig Date?", 
        'Unreportable', 
        "Housing Income Verification", 
        #'Income Verification Tester', 
        "Housing Posture of Case on Eligibility Date", 
        "Gen Case Index Number",'Case Number Tester',   
        "Housing Type Of Case", 
        "Housing Level of Service", 
        "Close Reason", 
        #'Close Tester',
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
        "Number of People 18 and Over","Number of People under 18",
        #"Over-18 Tester", 
        "Date of Birth",
        "Percentage of Poverty", 
        "Case Disposition", 
        "Housing Date Of Waiver Approval","Housing TRC HRA Waiver Categories","Waiver Tester", 
        "IOLA Outcome", 
        "Housing Signed DHCI Form", 
        "Income Types", 
        "Total Annual Income ", 
        "Housing Funding Note", 
        "Age at Intake",
        #"Over 62?",
        "Total Time For Case", 
        "Service Date", 
        "Caseworker Name", 
        "Retainer on File Compliance", 
        "Retainer on File", 
        "Case Involves Covid-19", 
        "Legal Problem Code", 
        "Agency",
        "Matter/Case ID#",
        "Housing Tab Assignment",
        #'Post-3/1 Limited Service Tester' 
         
 
        ]]       
         
        #FY22 Graveyard
        
        #Old Tab Assigner (not expansive) - was between level types and agency, prev commented out
        '''#Tab Assigner based on Primary Funding Code 
         
        def TabAssigner(PrimaryFundingCode): 
            if PrimaryFundingCode == "3011 TRC FJC Initiative" or PrimaryFundingCode == "3018 Tenant Rights Coalition (TRC)": 
                return "TRC" 
            elif PrimaryFundingCode == "3111 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3112 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3113 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3114 HRA-HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3115 HPLP-Homelessness Prevention Law Project" or PrimaryFundingCode == "3121 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3122 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3123 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3124 Universal Access to Counsel – (UAC)" or PrimaryFundingCode == "3125 Universal Access to Counsel – (UAC)": 
                return "UAHPLP" 
            else: 
                return "Other" 
                 
        df['Funding Code Sorter'] = df.apply(lambda x : TabAssigner(x['Primary Funding Code']),axis=1)''' 
        
        #label for pre-12/1 cases - was between the EDate Construct and the EDateConstructMaker
        '''#Label pre 12/1 eligible cases needing short data requirements 
        def PreDecOne (EligibilityDate):
            #commented out
            #if isinstance (EligibilityDate, int) == False: 
            #    return EligibilityDate 
            #elif EligibilityDate < 20211201: 
            #    return "Pre 12/1 Eligible Case"
            try: 
                if EligibilityDate < 20211201: 
                    return "Pre 12/1 Eligible Case" 
            except: 
                return EligibilityDate 
                 
                 
        df['Pre-12/1/21 Elig Date?'] = df.apply(lambda x: PreDecOne(x['EDate Construct']), axis=1) 
         
        #Post 12/1/21, cases must have Edate and Release 
        #-pre 12/1 LoS Advice, Blank LoS closed Advice, UAHPLP Brief, Blank LoS UAHPLP closed Brief, HfR closed Advice, UAHPLP HfR closed Brief,  don't need release '''
        
        #FY22 ReleaseAndEligTester - was between Today Construct and HousingTypeClean
        '''def FY22ReleaseAndEligTester(HRARelease,EligibilityDate, EliDC, TdC, LoS, Pre12, Close, PrimaryFunding): 
            UAHPLPFund = ['3121 Universal Access to Counsel – (UAC)','3122 Universal Access to Counsel – (UAC)','3123 Universal Access to Counsel – (UAC)','3124 Universal Access to Counsel – (UAC)','3125 Universal Access to Counsel – (UAC)','3111 HPLP-Homelessness Prevention Law Project','3112 HPLP-Homelessness Prevention Law Project','3113 HPLP-Homelessness Prevention Law Project','3114 HRA-HPLP-Homelessness Prevention Law Project','3115 HPLP-Homelessness Prevention Law Project'] 
            LoSEmpty = ["", "Hold For Review"] 
            if 20210701 > EliDC > 20000101: 
                return "Liz to check or remove old EDate Case" 
            elif 20000101 >= EliDC > 0: 
                return "Liz to review old EDate" 
            elif TdC < EliDC: 
                return "Liz to review future EDate" 
            elif HRARelease == 'Yes' and EligibilityDate != '': 
                return 'Reportable' 
            elif Pre12 == "Pre 12/1 Eligible Case": 
                if LoS == "Advice": 
                    return 'Reportable, Pre 12/1 advice' 
                elif LoS in LoSEmpty and Close == "A - Counsel and Advice": 
                    return "Reportable, Pre 12/1 Blank/Hold LoS Cl A" 
                elif LoS in LoSEmpty and PrimaryFunding in UAHPLPFund and Close == "B - Limited Action (Brief Service)": 
                    return "Reportable, Pre 12/1 Blank LoS UAHP Cl B" 
                elif LoS == "Brief Service" and PrimaryFunding in UAHPLPFund: 
                    return "Reportable, Pre 12/1 UAHP Brief" 
                else: 
                    return "Pre-12/1/21 Case Needs Release" 
            elif HRARelease == 'Yes' and EligibilityDate == '': 
                return "Needs Eligibility Date" 
            elif HRARelease != "Yes": 
                if EligibilityDate != '': 
                    return "Needs HRA Release" 
                else:    
                    return 'Needs Elig & Release' 
            else: 
                return 'Unknown Answer Category' 
        df['Unreportable'] = df.apply(lambda x: FY22ReleaseAndEligTester(x['HRA Release?'],x['HAL Eligibility Date'],x['EDate Construct'],x['TodayConstruct'],x['Housing Level of Service'],x['Pre-12/1/21 Elig Date?'],x['Close Reason'],x['Primary Funding Code']), axis=1)'''
        
        #Close Reason must match Level of Service in some way - between LoS and BuildingCaseClean
        #Does out of court advocacy need to be to a certain level? Full rep?***
        '''def CloseReasonClean (Close, LOS, PFC):
            TRC = ['3018 Tenant Rights Coalition (TRC)','3011 TRC FJC Initiative'] 
            if Close != "":
                if LOS == "Advice":
                    if Close == "A - Counsel and Advice":
                        return "Advice Fine"
                    else:
                        return "Needs alignment of Level of Service & Close Reason"
                elif LOS == "Brief Service":
                    if Close == "A - Counsel and Advice" and PFC in TRC:
                        return "TRC BA, fine"
                    elif Close == "B - Limited Action (Brief Service)" or Close == "F - Negotiated Settlement w/out Litigation":
                        return "Brief B or F, Fine"
                    else:
                        return "Brief service case needs close reason B or F"
                else:
                    return "Other cases untested, Liz to review"
            else:
                return "Close is Blank, fine"
        df['Close Tester'] = df.apply(lambda x:CloseReasonClean(x['Close Reason'],x['Housing Level of Service'],x['Primary Funding Code']), axis=1)''' 

        #Rent tester was between referral source and units clean, prev commented out, and ws 3rd in formats under secondary
        '''#monthly rent can't be 0 [sharene doesn't care if it's 0] 
        #df['Rent Tester'] = df.apply(lambda x: HousingToolBox.RentClean(x['Housing Total Monthly Rent']), axis=1) 
        df['Rent Tester'] = ''  '''
        
        '''     ws.conditional_format('AA2:AA100000',{'type': 'text', 
                                                 'criteria': 'containing', 
                                                 'value': "Needs", 
                                                 'format': bad_problem_format})'''
        
        #PA Tester was between Social Security # and CaseNumClean, two old PA Tester Versions
        '''#PA Tester (need to be correct format as well) 
                
        def PATesterClean (PANumber,SSTester): 
                         
            PANumber = str(PANumber) 
            LastCharacter = PANumber[-1:] 
            PenultimateCharater = PANumber[-2:-1] 
            SecondCharacter = PANumber [1:2] 
            if SSTester == '': 
                return 'Unnecessary due to SS#' 
            elif PANumber == '' or PANumber == 'None' or PANumber == 'NONE' or SecondCharacter == 'o': 
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
 
         
        df['PA # Tester'] = df.apply(lambda x: PATesterClean(x['Gen Pub Assist Case Number'],x['Social Security #']), axis=1) 
         
        #if PA # Tester is fine then SS# Tester doesn't matter 
         
        def SSDoubleTester(SSTester, PATester, PANum): 
            PANum = str(PANum) 
            SecondCharacter = PANum [1:2] 
            if PANum == '' or PANum == 'None' or PANum == 'NONE' or SecondCharacter == 'o' or SecondCharacter == 'n' or PANum == '000000000': 
                return SSTester 
            elif PATester == '' and len(PANum) >= 9: 
                return 'Unnecessary due to PA #' 
            else: 
                return SSTester 
         
        df['SS # Tester'] = df.apply(lambda x: SSDoubleTester(x['SS # Tester'],x['PA # Tester'],x['Gen Pub Assist Case Number']),axis=1)''' 
        
        #Calculate client's current age for over 62? - was between services rendered and activity tester
        '''def CurrentClientAge(BirthDate): 
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
 
             
        df['Over 62?'] = df.apply(lambda x: CurrentClientAge(x['Date of Birth']),axis = 1) '''
        
        #FY22 Waiver tester - was between Number of people over 18 and poverty tester
        #date of waiver approval & waiver categories - if there's something in one but not the other, then flag it. **in process, zip code? 
         
        '''def WaiverTester (WaiverType,WaiverDate,Poverty,RefSou,Type,LoS,PrimaryFunding,EDC): 
            UAHPLPFund = ['3121 Universal Access to Counsel – (UAC)','3122 Universal Access to Counsel – (UAC)','3123 Universal Access to Counsel – (UAC)','3124 Universal Access to Counsel – (UAC)','3125 Universal Access to Counsel – (UAC)','3111 HPLP-Homelessness Prevention Law Project','3112 HPLP-Homelessness Prevention Law Project','3113 HPLP-Homelessness Prevention Law Project','3114 HRA-HPLP-Homelessness Prevention Law Project','3115 HPLP-Homelessness Prevention Law Project']
            global HousingFundingCodes
            HousingFundingCodes= ['3121 Universal Access to Counsel – (UAC)','3122 Universal Access to Counsel – (UAC)','3123 Universal Access to Counsel – (UAC)','3124 Universal Access to Counsel – (UAC)','3125 Universal Access to Counsel – (UAC)','3111 HPLP-Homelessness Prevention Law Project','3112 HPLP-Homelessness Prevention Law Project','3113 HPLP-Homelessness Prevention Law Project','3114 HRA-HPLP-Homelessness Prevention Law Project','3115 HPLP-Homelessness Prevention Law Project','3018 Tenant Rights Coalition (TRC)','3011 TRC FJC Initiative'] 
            if Poverty < 201: 
                if WaiverType == "" and WaiverDate == "":
                    return "Income Eligible, <201%"
                else:
                    return "Needs waiver fields cleared, income eligible case does not need waiver"
            elif PrimaryFunding == "3011 TRC FJC Initiative":
                if WaiverDate == "" and WaiverType == "":
                    return "Needs waiver type, 'FJC Waiver' and waiver date '11/28/2016'"
                elif WaiverDate != "11/28/2016": 
                    return "Needs waiver date '11/28/2016'"
                elif WaiverType == "":
                    return "Needs waiver type 'FJC Waver'"
                else:
                    return "FJC waiver, fine"
            elif PrimaryFunding != "3011 TRC FJC Initative" and WaiverType == "FJC Waiver":
                return "Needs change, only cases referred by FJC (3011) can have FJC waiver"
            elif WaiverType != "" and WaiverDate != "": 
                return "Already has waiver"
            elif WaiverDate != "" and WaiverType == "": 
                return "Needs Waiver Type" 
            elif WaiverType != "" and WaiverDate == "":
                return "Needs Waiver Date"
            elif PrimaryFunding not in HousingFundingCodes:
                return "Not housing case, waiver not needed at this time"
            elif "Needs" in LoS:
                return "Missing Level of Service, waiver need unknown"
            elif "Needs" in Type:
                return "Missing Type of Case, waiver need unknown"
            elif PrimaryFunding in HousingFundingCodes:
                if LoS == "Advice":
                    return "Housing Advice, no waiver needed"
                elif PrimaryFunding in UAHPLPFund and LoS == "Brief Service":
                    return "UAHPLP Brief, no waiver needed"
                elif Type == "Illegal Lockout":
                    return "Needs update, ILO has categorical income waiver with date '09/20/2021'"
                elif 20210630 < EDC < 20220318:
                    if LoS.startswith('Representation') == True and Type in evictiontypes: 
                        return "Needs update, pre 3/18/22 full rep evic has categorical income waiver w date '09/20/2021'" 
                    else: 
                        return "pre 3/18 non Full Rep eviction case, may need waiver requested"
                elif LoS.startswith('Representation') == True and Type in evictiontypes:
                    return "if Edate pre 3/18/22, has categorical waiver, otherwise may need waiver requested"
                else:
                    return "Possible waiver request needed for non-eviction / non-full rep eviction cases"
            else: 
                return "Possible waiver request needed - Liz to check" 
        df['Waiver Tester'] = df.apply(lambda x: WaiverTester(x['Housing TRC HRA Waiver Categories'],x['Housing Date Of Waiver Approval'],x['Percentage of Poverty'],x['Referral Source'],x['Housing Type Of Case'],x['Housing Level of Service'],x['Primary Funding Code'],x['EDate Construct']), axis = 1) '''
        
        #FY22 Blue highlighting for low info needs cases, started after problem_format for 'Tester'
        '''     ws.conditional_format('C1:BO1',{'type': 'text', 
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
                                                 'format': medium_problem_format})'''
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
                Liz_check_format = workbook.add_format({'bg_color':'#4397EC'}) 
                medium_problem_format = workbook.add_format({'bg_color':'cyan'}) 
                ws.set_column('A:A',20,link_format) 
                ws.set_column('B:B',16) 
                ws.set_column('D:ZZ',25) 
                ws.set_column('C:C',18) 
                ws.autofilter('B1:CG1') 
                ws.freeze_panes(1, 2) 
                C2BOFullRange='C2:BO'+str(dict_df[i].shape[0]+1) 
                print("BORowRange is "+ str(C2BOFullRange)) 
                ws.conditional_format(C2BOFullRange,{'type': 'text', 
                                                 'criteria': 'containing', 
                                                 'value': 'No Release - Remove Elig Date', 
                                                 'format': bad_problem_format}) 
                ws.conditional_format('AA2:AA100000',{'type': 'text', 
                                                 'criteria': 'containing', 
                                                 'value': "secondary", 
                                                 'format': bad_problem_format}) 
                ws.conditional_format(C2BOFullRange,{'type': 'text', 
                                                 'criteria': 'containing', 
                                                 'value': "Liz to", 
                                                 'format': Liz_check_format})
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
                                                 'value': 'Social', 
                                                 'format': medium_problem_format}) 
                ws.conditional_format('C1:BO1',{'type': 'text', 
                                                 'criteria': 'containing', 
                                                 'value': 'Index', 
                                                 'format': medium_problem_format})                                                 
                                                  
                                                               
            writer.save() 
         
        output_filename = f.filename 
         
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename) 
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename) 
 
    return ''' 
    <!doctype html> 
    <title>All Housing Simpler</title> 
    <link rel="stylesheet" href="/static/css/main.css">   
    <h1>All Housing Simpler</h1> 
    <form action="" method=post enctype=multipart/form-data> 
    <p><input type=file name=file><input type=submit value=Clean!> 
    </form> 
    <h3>Instructions:</h3> 
    <ul type="disc"> 
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2452" target="_blank">All Housing Python Report</a>.</li> 
     
     
    </br> 
    <a href="/">Home</a> 
    ''' 
