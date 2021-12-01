#HOUSING CODE

#List of proceeding types that constitute an eviction case
evictionproceedings = ['HO','NP','IL','TT','EA','EJ']
#Housing Type of Case Eviction-Types:
evictiontypes = ['Holdover','Non-payment','Illegal Lockout','NYCHA Housing Termination']
#Highest Level of Service Reps
leveltypes = ['Representation - State Court','Representation - Federal Court']


#Functions to Help with TRC/UAHPLP Cleanup


def NonHousingTester (ProblemCode, EligConstruct):
    if EligConstruct != '':
        if ProblemCode.startswith('6') != True and EligConstruct > 20200930:
            return 'Needs Review'
        else:
            return ''
    else:
        return ''

#Has to have an HRA Release
def HRAReleaseClean (HRARelease,EligibilityDate):
    if HRARelease == 'No' and EligibilityDate != '':
        return 'No Release - Remove Elig Date'
    elif HRARelease == '' and EligibilityDate != '':
        return 'No Release - Remove Elig Date'
    elif HRARelease == 'Yes':
        return ''
    else:
        return 'Needs HRA Release'
        
def ReleaseAndEligTester(HRARelease,EligibilityDate):
    if HRARelease == 'Yes' and EligibilityDate != '':
        return ''
    elif HRARelease == 'Yes' and EligibilityDate == '':
        return "Needs Eligibility Date"
    elif HRARelease == 'No' or HRARelease == '':
        if EligibilityDate != '':
            return "Needs HRA Release"
        else:   
            return 'Needs Elig & Release'
    else:
        return 'Needs Elig & Release'

#Has to have a Housing Type of Case
def HousingTypeClean (HousingType):
    if HousingType == '':
        return 'Needs Housing Type of Case'
    else:
        return ''
 
#Has to have a Housing Level of Service - and if DHCR case has to be admin proceeding or rep - admin agency. 
#TRCCovidClean
#...and? 
def HousingLevelClean (HousingLevel,HousingType):
    if HousingLevel == '':
        return 'Needs Level of Service'
    else:
        return '' 


#Has to say whether or not it's a building case        
#TRCCovidClean
#...and?
def BuildingCaseClean (BuildingCase):
    if BuildingCase == '':
        return 'Needs Building Case Answer'
    if BuildingCase == 'Prefer Not To Answer':
        return 'Needs Building Case Answer'
    else:
        return ''
  
#Referral Source Can't Be Blank        
def ReferralClean (Referral,FundingSource):
    if Referral == '':
        return 'Needs Referral Source'
    elif FundingSource == '3011 TRC FJC Initiative' and Referral != 'FJC Housing Intake':
        return 'Must be FJC'
    else:
        return ''

#rent can't be 0
def RentClean (Rent):
    if Rent == 0:
        return 'Needs Rent Amount'
    else:
        return ''

#number of units in building can't be 0 or written with letters        
def UnitsClean (Units):
    
    if Units == 0:
        return 'Needs Units'
    Units = str(Units)
    if any(c.isalpha() for c in Units) == True:
        return 'Needs To Be Number'
    elif Units == "0":
        return 'Needs Units'
    else:
        return ''
        
#Housing form of regulation can't be blank      
def RegulationClean (Regulation):
    if Regulation == '':
        return 'Needs Form of Regulation'
    else:
        return ''


#Housing subsidy can't be blank (can be none)
def SubsidyClean (Subsidy):
    if Subsidy == '':
        return 'Needs Type of Subsidy'
    else:
        return ''


#Years in Apartment Can't be 0 (can be -1)
def YearsClean (Years):
    if Years == 0:
        return 'Needs Years In Apartment'
    elif Years < -1 or Years > 99:
        return 'Needs Valid Number'
    else:
        return ''

#Language Can't be Blank
        
def LanguageClean (Language):
    if Language == '' or Language == 'Unknown':
        return 'Needs Language'
    else:
        return ''

#Housing Posture of Case can't be blank if there is an eligibility date
        
def PostureClean (Posture,EligibilityDate,Type,Level):
    if Type in evictiontypes:
        if Level.startswith("Rep") == True:
        
            if EligibilityDate  != '' and Posture == '':
                return 'Needs Posture of Case'
            else:
                return ''
        else:
            return ''
    else:
        return ''

#TRC Housing Income Verification can't be blank or none and other stuff with kids and poverty level and you just give up if it's closed
def IncomeVerificationClean (IncomeVerification, Children, PovertyPercent, Disposition, LevelOfService):
    if LevelOfService == 'Advice':
        return ''
    elif Children > 0 and PovertyPercent <= 200.99 and IncomeVerification == '':
        return 'Must Have DHCI or PA#'
    elif Children > 0 and PovertyPercent <= 200.99 and IncomeVerification == 'None':
        return 'Must Have DHCI or PA#'
    elif Disposition == 'Closed' and IncomeVerification =='None':
        return ''
    elif IncomeVerification == '':
        return 'Needs Income Verification'
    else:
        return ''
        
#UAC Housing Income Verification can't be blank or none and other stuff with kids and poverty level and you just give up if it's closed (after 6 months)
def UACIncomeVerificationClean (IncomeVerification, Children, PovertyPercent, Disposition, DateClosed,TodayDate):
    if Children > 0 and PovertyPercent <= 200.99 and IncomeVerification == '':
        return 'Must Have DHCI or PA#'
    elif Children > 0 and PovertyPercent <= 200.99 and IncomeVerification == 'None':
        return 'Must Have DHCI or PA#'
    
    
    elif Disposition == 'Closed' and IncomeVerification =='None' and TodayDate-DateClosed >= 600 and TodayDate-DateClosed <= 8870:
        return ''
    elif Disposition == 'Closed' and IncomeVerification =='None' and TodayDate-DateClosed >= 9400:
        return ''
    elif Disposition == 'Closed' and IncomeVerification =='' and TodayDate-DateClosed >= 600 and TodayDate-DateClosed <= 8870:
        return ''
    elif Disposition == 'Closed' and IncomeVerification =='' and TodayDate-DateClosed >= 9400:
        return ''
    elif IncomeVerification == '':
        return 'Needs Income Verification'
    elif IncomeVerification == 'None':
        return 'Needs Income Verification'
    else:
        return ''

#PA Tester (need to be correct format as well)
def PATesterClean (PANumber):
                
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

#Test if case number is correct format (don't need one if it's brief, advice, or out-of-court)
def CaseNumClean (CaseNum,Level):
    CaseNum = str(CaseNum)
    First3 = CaseNum[0:3]
    ThirdFromEnd = CaseNum[-3:-2]
    SecondFromEnd = CaseNum[-2:-1]
    First6 = CaseNum[0:6]
    First2 = CaseNum[0:2]
    
    if Level == 'Advice' or Level == 'Brief Service' or Level == 'Out-of-Court Advocacy' or Level == 'Hold For Review' or Level.startswith('UAC') == True:
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
    #NYCHA Housing Termination Cases -  "904607-CR-2019"
    elif len(CaseNum) == 14 and str.isdigit(First6) == True:
        return ''
    else:
        return "Needs Correct Case # Format"

#Test if social security number is correct format (or ignore it if there's a valid PA number)
def SSNumClean (CaseNum, PANumber):
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

        
#Test Housing Activity Indicator - can't be blank for closed cases that are full rep state or full rep federal(housing level of service) and eviction cases(housing type of case: non-payment holdover illegal lockout nycha housing termination)    

def ActivityTesterClean(HousingActivity,Disposition,Level,Type):
    if Disposition == 'Closed' and Level in leveltypes and Type in evictiontypes and HousingActivity == '':
        return 'Needs Activity Indicator'
    else:
        return ''


#Test Housing Services Rendered - can't be blank for closed cases that are full rep state or full rep federal(housing level of service)
        
def ServicesTesterClean(HousingServices,Disposition,Level,Type):
    if Disposition == 'Closed' and Level in leveltypes and HousingServices == '':
        return 'Needs Services Rendered'
    elif Level == 'Representation - Admin. Agency' and Disposition == 'Closed' and HousingServices == '':
        return 'Needs Services Rendered'
    elif Level.startswith('UAC') == True and Disposition == 'Closed' and HousingServices == '':
        return 'Needs Services Rendered'
    else:
        return ''

#Outcome Tester - needs outcome and date for closed  eviction cases that are full rep at state or federal level (not admin)
        
def TRCOutcomeTesterClean (Disposition,Outcome,OutcomeDate,Level,Type):
    if Disposition == 'Closed' and Level in leveltypes and Type in evictiontypes and Outcome == '' and OutcomeDate == '':
        return 'Needs Outcome & Date'
    elif Disposition == 'Closed' and Level in leveltypes and Type in evictiontypes and Outcome == '':
        return 'Needs Outcome'
    elif Disposition == 'Closed' and Level in leveltypes and Type in evictiontypes and OutcomeDate == '':
        return 'Needs Outcome Date'
    else:
        return ''



#Outcome Tester - IF CLOSED OR IF OVER 6 Months Old (since eligiblity date) needs outcome and date for eviction cases that are full rep at state or federal level (not admin)
        
def UAHPLPOutcomeTesterClean (Disposition,Outcome,OutcomeDate,Level,Type,EligDate,TodayDate):
    if EligDate == "":
        return ""
    elif TodayDate-EligDate >= 600 and TodayDate-EligDate <= 8870:
        if Level in leveltypes and Type in evictiontypes and Outcome == '' and OutcomeDate == '':
            return 'Needs Outcome & Date'
        elif Level in leveltypes and Type in evictiontypes and Outcome == '':
            return 'Needs Outcome'
        elif Level in leveltypes and Type in evictiontypes and OutcomeDate == '':
            return 'Needs Outcome Date'
        else:
            return ''
    elif TodayDate-EligDate >= 9400:
        if Level in leveltypes and Type in evictiontypes and Outcome == '' and OutcomeDate == '':
            return 'Needs Outcome & Date'
        elif Level in leveltypes and Type in evictiontypes and Outcome == '':
            return 'Needs Outcome'
        elif Level in leveltypes and Type in evictiontypes and OutcomeDate == '':
            return 'Needs Outcome Date'
        else:
            return ''
    elif Disposition == 'Closed' :
        if Level in leveltypes and Type in evictiontypes and Outcome == '' and OutcomeDate == '':
            return 'Needs Outcome & Date'
        elif Level in leveltypes and Type in evictiontypes and Outcome == '':
            return 'Needs Outcome'
        elif Level in leveltypes and Type in evictiontypes and OutcomeDate == '':
            return 'Needs Outcome Date'
        else:
            return ''
    else:
        return ''


#Functions to be used in TRC/UAHPLP Housing Report Prep 

#Translation based on HRA Specs
def TRCProceedingType(TypeOfCase,LegalProblemCode,LevelOfService,EligDate):
    if EligDate < 20201001 and LegalProblemCode.startswith("0") == True:
        return "CON"
    elif EligDate < 20201001 and LegalProblemCode.startswith("3") == True:
        return "FAM"
    elif EligDate < 20201001 and LegalProblemCode.startswith("5") == True:
        return "HEA"
    elif EligDate < 20201001 and LegalProblemCode.startswith("7") == True:
        return "BEN"
    elif TypeOfCase == "HP Action":
        return "HP"
    elif TypeOfCase == "Affirmative Litigation Supreme":
        return "OS"
    elif TypeOfCase == "Holdover":
        return "HO"
    elif TypeOfCase == "No Case" or TypeOfCase == "Non-Litigation Advocacy" or TypeOfCase == "Tenant Rights"  or TypeOfCase == "Rent Strike":
        return "OO"
    elif TypeOfCase == "Non-payment":
        return "NP"
    elif TypeOfCase == "Section 8 Administrative Proceeding" or TypeOfCase == "Sec. 8 Termination" or TypeOfCase == "Section 8 Grievance" or TypeOfCase == "Section 8 HQS" or TypeOfCase == "Section 8 other" or TypeOfCase == "Section 8 share":
        return "S8"
    elif TypeOfCase == "Illegal Lockout":
        return "IL"
    elif TypeOfCase == "NYCHA Termination of Tenancy" or TypeOfCase == "NYCHA Housing Termination":
        return "TT"
    elif TypeOfCase == "Dispositive Eviction Appeal":
        return "EA"
    elif TypeOfCase == "Ejectment Action":
        return "EJ"
    elif TypeOfCase == "DHCR Administrative Action" or TypeOfCase == "DHCR Proceeding":
        return "DA"
    elif TypeOfCase == "7A Proceeding":
        return "7A"
    elif TypeOfCase == "Article 78":
        return "78"
    elif TypeOfCase == "Affirmative Litigation Federal" or TypeOfCase == "Appeal Federal":
        return "FC"
    elif TypeOfCase == "HRA Fair Hearing" or TypeOfCase == "Human Rights Complaint" or TypeOfCase == "Mitchell-Lama RFM"or TypeOfCase == "Mitchell-Lama Termination"or TypeOfCase == "NYCHA Housing Grievance"or TypeOfCase == "NYCHA RFM"or TypeOfCase == "Other Administrative Proceeding" or TypeOfCase == "PA Issue: City FEPS/SEPS" or TypeOfCase == "PA Issue: Budgeting" or TypeOfCase == "PA Issue: FEPS" or TypeOfCase == "PA Issue: LINC" or TypeOfCase == "PA Issue: Other" or TypeOfCase == "PA Issue: RAU" or TypeOfCase == "PA Issue: Underpayment" or TypeOfCase == "SCRIE/DRIE" or TypeOfCase == "Certificate of No Harassment Case":
        return "OA"
    elif TypeOfCase == "Affirmative Litigation Supreme" or TypeOfCase == "Public Nuisance" or TypeOfCase == "Appeal Supreme" or TypeOfCase == "Other Civil Court":
        return "OS"
    elif TypeOfCase == "Appeal-Appellate Division" or TypeOfCase == "Appeal-Appellate Term":
        return "EA"
    elif TypeOfCase == "Illegal Lockout":
        return "IL"
    else:
        return "Needs Review"
 
def UACProceedingType(TypeOfCase,LegalProblemCode,CloseReason,LevelOfService):
    

    if LegalProblemCode.startswith("0") == True:
        return "CON"
    elif LegalProblemCode.startswith("3") == True:
        return "FAM"
    elif LegalProblemCode.startswith("5") == True:
        return "HEA"
    elif LegalProblemCode.startswith("7") == True:
        return "BEN"
    
    elif TypeOfCase == "HP Action":
        return "HP"
    elif TypeOfCase == "Affirmative Litigation Supreme":
        return "OS"
    elif TypeOfCase == "Holdover":
        return "HO"
    elif TypeOfCase == "No Case" or TypeOfCase == "Non-Litigation Advocacy" or TypeOfCase == "Tenant Rights"  or TypeOfCase == "Rent Strike":
        return "EVC"
    elif TypeOfCase == "Non-payment":
        return "NP"
    elif TypeOfCase == "Section 8 Administrative Proceeding" or TypeOfCase == "Sec. 8 Termination" or TypeOfCase == "Section 8 Grievance" or TypeOfCase == "Section 8 HQS" or TypeOfCase == "Section 8 other" or TypeOfCase == "Section 8 share":
        return "S8"
    elif TypeOfCase == "Illegal Lockout":
        return "IL"
    elif TypeOfCase == "NYCHA Termination of Tenancy":
        return "TT"
    elif TypeOfCase == "Dispositive Eviction Appeal":
        return "EA"
    elif TypeOfCase == "Ejectment Action":
        return "EJ"
    elif TypeOfCase == "DHCR Administrative Action" or TypeOfCase == "DHCR Proceeding":
        return "DA"
    elif TypeOfCase == "7A Proceeding":
        return "7A"
    elif TypeOfCase == "Article 78":
        return "78"
    elif TypeOfCase == "Affirmative Litigation Federal" or TypeOfCase == "Appeal Federal":
        return "FC"
    elif TypeOfCase == "HRA Fair Hearing" or TypeOfCase == "Human Rights Complaint" or TypeOfCase == "Mitchell-Lama RFM"or TypeOfCase == "Mitchell-Lama Termination"or TypeOfCase == "NYCHA Housing Grievance"or TypeOfCase == "NYCHA RFM"or TypeOfCase == "Other Administrative Proceeding" or TypeOfCase == "PA Issue: City FEPS/SEPS" or TypeOfCase == "PA Issue: Budgeting" or TypeOfCase == "PA Issue: FEPS" or TypeOfCase == "PA Issue: LINC" or TypeOfCase == "PA Issue: Other" or TypeOfCase == "PA Issue: RAU" or TypeOfCase == "PA Issue: Underpayment" or TypeOfCase == "SCRIE/DRIE" or TypeOfCase == "Certificate of No Harassment Case":
        return "OA"
    elif TypeOfCase == "Affirmative Litigation Supreme" or TypeOfCase == "Public Nuisance" or TypeOfCase == "Appeal Supreme" or TypeOfCase == "Other Civil Court":
        return "OS"
    elif TypeOfCase == "Appeal-Appellate Division" or TypeOfCase == "Appeal-Appellate Term":
        return "EA"
    elif TypeOfCase == "Illegal Lockout":
        return "IL"
    else:
        return ""

 


#List of UAC (RTC) Zip Codes:
UACZipCodes = ['11207','11216','11221','11225','11226','10453','10457','10462','10467','10468','10025','10026','10027','10029','10031','10034','10302','10303','10310','10314','11373','11385','11433','11434','11691']

#Case posture on eligibility date (on trial, no stipulation etc.) - transform them into the HRA initials
def PostureOnEligibility(Posture):
    splitpostureeligibilitylist = Posture.split( ", ")
    recombinedposturelist = list()
    for x in splitpostureeligibilitylist:
        if x == "No Stipulation; No Judgment":
            x = "NSNJ"
            recombinedposturelist.append(x)
        elif x == "On for Trial":
            x = "OFT"
            recombinedposturelist.append(x)
        elif x == "Post-Stipulation":
            x = "PSNJ"
            recombinedposturelist.append(x)
        elif x == "Tenant in Possession-Judgment Due to Default":
            x = "PJD"
            recombinedposturelist.append(x)
        elif x == "Tenant in Possession-Judgment Due to Other":
            x = "PJO"
            recombinedposturelist.append(x)
        elif x == "Tenant Out of Possession":
            x = "PJP"
            recombinedposturelist.append(x)
 
    return "; ".join(recombinedposturelist)
 
#TRC Level of Service becomes Service type - lots of level of service in LS, mapped to Advice Only, Pre-Lit Strategies (brief service, out of court advocacy, hold for review), and Full Rep (mapping to be confirmed)
def TRCServiceType(LevelOfService,LegalProblemCode,FundingCode, Referral, HRARelease):
    if LevelOfService == "Advice":
        return "Advice Only"
    #elif LegalProblemCode.startswith("3") == True or LegalProblemCode.startswith("5") == True or LegalProblemCode.startswith("7") == True:
    #    return "Advice Only"
    elif FundingCode.startswith('3011') == True and HRARelease != 'Yes': 
        return "Advice Only"
    elif LevelOfService == "Brief Service" or LevelOfService == "Out-of-Court Advocacy" or LevelOfService == "Hold For Review":
        return "Pre-Litigation Strategies"
    elif LevelOfService == "Representation - Admin. Agency" or LevelOfService == "Representation-EOIR" or LevelOfService == "Representation - Federal Court" or LevelOfService == "Representation - State Court":
        return "Full Rep"

#UAC Service Type: 
def UACServiceType(LevelOfService,UAorNonUA,CloseReason,LegalProblemCode):
    if CloseReason.startswith(("A","B")) == True:
        if UAorNonUA == "UA":
            return "Brief Legal Assistance"
        elif UAorNonUA == "Non-UA":
            return "Advice Only"
    elif LegalProblemCode.startswith("0") == True  or LegalProblemCode.startswith("3") == True or LegalProblemCode.startswith("5") == True or LegalProblemCode.startswith("7") == True:
        if UAorNonUA == "UA":
            return "Brief Legal Assistance"
        elif UAorNonUA == "Non-UA":
            return "Advice Only"
    elif CloseReason.startswith(("F","G","H","IA","IB","L")) == True:
        return "Full Rep"
    elif LevelOfService == "Advice" or LevelOfService == "Brief Service" or LevelOfService == "Out-of-Court Advocacy":
        if UAorNonUA == "UA":
            return "Brief Legal Assistance"
        elif UAorNonUA == "Non-UA":
            return "Advice Only"
    elif LevelOfService == "Hold For Review":
        return "Hold For Review"
    elif LevelOfService == "Representation - Admin. Agency" or LevelOfService == "Representation-EOIR" or LevelOfService == "Representation - Federal Court" or LevelOfService == "Representation - State Court" or LevelOfService.startswith('UAC') == True:
        return "Full Rep"
 
#Housing Regulation Type: mapping down - we have way more categories, rent regulated, market rate, or other (mapping to be confirmed). can't be blank     
def HousingType(HousingType):
    if HousingType == "Low Income Tax Credit" or HousingType == "Tenant-interim-lease" or HousingType == "Mitchell-Lama" or HousingType == "Rent Stabilized" or HousingType == "Rent Controlled" or HousingType == "Project-based Sec. 8" or HousingType == "Other Subsidized Housing" or HousingType == "Supportive Housing":
        return "Rent-Regulated"
    elif HousingType == "HDFC" or HousingType == "Unknown":
        return "Other"
    elif HousingType == "Unregulated" or HousingType == "Unregulated - Sublet" or HousingType == "Unregulated - Co-Op" or HousingType == "Unregulated - Other" or HousingType == "Market Rate – Sublet" or HousingType == "Market Rate – Co-Op" or HousingType == "Market Rate – Other":
        return "Market Rate"
    elif HousingType == "Public Housing" or HousingType == "Public Housing/NYCHA" or HousingType == "NYCHA/NYCHA":
        return "NYCHA"
        
#Referrals need to be one of their specific categories        
def ReferralMap(ReferralSource):

    if ReferralSource == "HRA ELS Part F Brooklyn" or ReferralSource == "HRA ELS (Assigned Counsel)" or ReferralSource == "HRA" or ReferralSource == "Documented Documented HRA Referral Referral":
        return "Documented HRA Referral"
        
    elif ReferralSource == "Court Referral-NON HRA" or ReferralSource == "Court": 
        return "Documented Judicial Referral"
        
    elif ReferralSource == "Tenant Support Unit":
        return "Public Engagement Unit/Tenant Support Unit"
        
    elif ReferralSource == "FJC Housing Intake":
        return "Family Justice Center"
    
    elif ReferralSource == "3-1-1" or ReferralSource == "ADP Hotline" or ReferralSource == "Community Organization" or ReferralSource == "Elected Official" or ReferralSource == "Foreclosure" or ReferralSource == "Friends/Family" or ReferralSource == "Home base" or ReferralSource == "In-House" or ReferralSource == "Other City Agency" or ReferralSource == "Outreach" or ReferralSource == "Returning Client" or ReferralSource == "School" or ReferralSource == "Self-referred" or ReferralSource == "Word of mouth" or ReferralSource == "Legal Services" or ReferralSource == "Other" or ReferralSource == "":
        return "Other"
        
#Housing Outcomes needs mapping for HRA   
def Outcome(HousingOutcome):
    if HousingOutcome == "Client Allowed to Remain in Residence":
        return "Remain"
    elif HousingOutcome == "Client Required to be Displaced from Residence":
        return "Displaced"
    elif HousingOutcome == "Client Discharged Attorney":
        return "Discharged"
    elif HousingOutcome == "Attorney Withdrew":
        return "Withdrew"
        
#Outcome related things that need mapping
def ServicesRendered(ServicesRendered):
    splitserviceslist = ServicesRendered.split(", ")
    recombinedserviceslist = list()
    
    for x in splitserviceslist:
        if x == "Secured Rent Abatement":
            x = "Abatement"
        elif x == "Secured Order or Agreement for Repairs in Apartment/Building":
            x = "Repairs"
        elif x == "Returned Unit to Rent Regulation":
            x = "Regulation"
        elif x == "Obtained Renewal of Lease":
            x = "Renewal"
        elif x == "Obtain Ongoing Rent Subsidy":
            x = "Subsidy"
        elif x == "Client Security Deposit Returned":
            x = "Deposit"
        elif x == "Case Discontinued/Dismissed/Landlord Fails to Prosecute":
            x = "Discontinued"
        elif x == "Case Resolved without Judgment of Eviction Against Client":
            x = "Resolved"
        elif x == "Secured 6 Months or Longer in Residence":
            x = "6months"
        elif x == "Obtained Succession Rights to Residence":
            x = "Succession"
        elif x == "Obtained Negotiated Buyout":
            x = "Buyout"
        elif x == "Restored Access to Personal Property":
            x = "Property"
        elif x == "Overcame Housing Discrimination":
            x = "Discrimination"
        elif x == "Provided Housing-related Consumer Debt Legal Assistance":
            x = "Debt"
        recombinedserviceslist.append(x)
    return "; ".join(recombinedserviceslist)
        
#Mapped to what HRA wants - some of the options are in LegalServer,
def Activities(Activity):
    splitactivitieslist = Activity.split(", ")
    recombinedactivitieslist = list()  
    for x in splitactivitieslist:
        if x == "Counsel Assisted in Filing or Refiling of Answer":
            x = "Answer"
        elif x == "Filed/Argued/Supplemented Dispositive or other Substantive Motion":
            x = "Motion"
        elif x == "Filed for an Emergency Order to Show Cause":
            x = "OSC"
        elif x == "Conducted Traverse Hearing":
            x = "Traverse"
        elif x == "Conducted Evidentiary Hearing":
            x = "Evidentiary"
        elif x == "Commenced Trial":
            x = "Trial"
        elif x == "Filed Appeal":
            x = "Appeal"        
        recombinedactivitieslist.append(x)
    return "; ".join(recombinedactivitieslist)
        
#Subsidy type - if it's not in the HRA list, it has to be 'none' (other is not valid) - they want a smaller list than we record. (mapping to be confirmed)        
def SubsidyType(SubsidyType):
    splitsubsidytypelist = SubsidyType.split(", ")
    recombinedsubsidytypelist = list()
    for x in splitsubsidytypelist:
        if x == "LINC" or x == "HOMETBRA" or x == "FEPS" or x == "SEPS" or x == "City FEPS" or x == "HASA" or x == "Pathways Home" or x == "SOTA" or x == "City HRA Subsidy":
            x = "HRA Subsidy"
        elif x == "HUD VASH":
            x = "Section 8"
        elif x == "ACS Housing Subsidy":
            x = "ACS Subsidy"
        recombinedsubsidytypelist.append(x)
        
    if "; ".join(recombinedsubsidytypelist) == "":
        return "None"
    else:
        return "; ".join(recombinedsubsidytypelist)

#Does Client have an eligibility date prior to March 1st, 2020?
def PreThreeOne(EligibilityDate):
    if isinstance(EligibilityDate, int) == False:
        return "No"
    elif EligibilityDate < 20200301:
        return "Yes"
    elif EligibilityDate >= 20200301:
        return "No"
        
#Does Client have an eligibility date after Dec 1st, 2021?
def PostTwelveOne(EligibilityDate):
    if isinstance(EligibilityDate, int) == False:
        return "Undetermined"
    elif EligibilityDate > 20211201:
        return "Yes"
    elif EligibilityDate <= 20211130:
        return "No"

def NeedsRedactingTester(LevelOfService, PreThreeOne,FundingCodeSorter):
    if LevelOfService.startswith("Advice") == True and PreThreeOne == "No":
        return "Needs Redacting"
    elif LevelOfService.startswith("Brief") == True and PreThreeOne == "No" and FundingCodeSorter == "UAHPLP":
        return "Needs Redacting"
    else:
        return ""
        
      
def TRCRedactForCovid(Edate, LevelOfService, ToRedact, PrimaryFunding, HRARelease):
            LevelOfService = str(LevelOfService)
            if Edate >= 20211201:
                return ToRedact
            elif LevelOfService.startswith("Advice") == True and ToRedact != "":
                if PrimaryFunding == "3011 TRC FJC Initiative" and HRARelease == "Yes":
                    return ToRedact
                else:
                    return ""
            else:   
                return ToRedact
                
def AllHousingRedactForCovid(LevelOfService, PreThreeOne, ToRedact,FundingCodeSorter):
            LevelOfService = str(LevelOfService)
            if LevelOfService.startswith("Advice") == True and PreThreeOne == "No" and ToRedact != "":
                return ""
            elif LevelOfService.startswith("Brief") == True and PreThreeOne == "No" and ToRedact != "" and FundingCodeSorter == "UAHPLP":
                return ""
            else:   
                return ToRedact
      
def RedactForCovid(LevelOfService, PreThreeOne, ToRedact):
            LevelOfService = str(LevelOfService)
            if LevelOfService.startswith("Advice") == True and PreThreeOne == "No" and ToRedact != "":
                return ""
            elif LevelOfService.startswith("Brief") == True and PreThreeOne == "No" and ToRedact != "":
                return ""
            else:   
                return ToRedact
            
        
def NoReleaseRedactForCovid(LevelOfService, PreThreeOne, ToRedact,Release):
            
            LevelOfService = str(LevelOfService)
            if Release == "Yes":
                return ToRedact
            elif LevelOfService.startswith("Advice") == True and PreThreeOne == "No" and ToRedact != "":
                return ""
            elif LevelOfService.startswith("Brief") == True and PreThreeOne == "No" and ToRedact != "":
                return ""
            else:   
                return ToRedact
            
      