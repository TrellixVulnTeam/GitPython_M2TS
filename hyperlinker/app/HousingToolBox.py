#General Purpose Functions to be used in LSNYC Report Prep 
#if LevelOfService.startswith("A") == True or LevelOfService.startswith("B") == True:
#Translation based on HRA Specs
def TRCProceedingType(TypeOfCase,LegalProblemCode,LevelOfService):
    if LegalProblemCode.startswith("0") == True and LevelOfService.startswith("A") == True:
        return "CON"
    elif LegalProblemCode.startswith("3") == True and LevelOfService.startswith("A") == True:
        return "FAM"
    elif LegalProblemCode.startswith("5") == True and LevelOfService.startswith("A") == True:
        return "HEA"
    elif LegalProblemCode.startswith("7") == True and LevelOfService.startswith("A") == True:
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
        return "Needs Review"
 
def UACProceedingType(TypeOfCase,LegalProblemCode,CloseReason,LevelOfService):
    
    if CloseReason.startswith("A") == True or CloseReason.startswith("B") == True or LevelOfService.startswith("A") == True or LevelOfService.startswith("B") == True or LevelOfService.startswith("O") == True:
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
        return "OO"
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
        return "Needs Review"
 
#List of proceeding types that constitute an eviction case
evictionproceedings = ['HO','NP','IL','TT','EA','EJ']

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
def TRCServiceType(LevelOfService):
    if LevelOfService == "Advice":
        return "Advice Only"
    elif LevelOfService == "Brief Service" or LevelOfService == "Out-of-Court Advocacy" or LevelOfService == "Hold For Review":
        return "Pre-Litigation Strategies"
    elif LevelOfService == "Representation - Admin. Agency" or LevelOfService == "Representation-EOIR" or LevelOfService == "Representation - Federal Court" or LevelOfService == "Representation - State Court":
        return "Full Rep"

#UAC Service Type: 
def UACServiceType(LevelOfService,UAorNonUA,CloseReason):
    if CloseReason.startswith(("A","B")) == True:
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
    elif LevelOfService == "Representation - Admin. Agency" or LevelOfService == "Representation-EOIR" or LevelOfService == "Representation - Federal Court" or LevelOfService == "Representation - State Court":
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
        
    