#General Purpose Functions to be used in LSNYC Report Prep 

#Translation based on HRA Specs
def ProceedingType(TypeOfCase):
    if TypeOfCase == "HP Action":
        return "HP"
    elif TypeOfCase == "Affirmative Litigation Supreme":
        return "OS"
    elif TypeOfCase == "Holdover":
        return "HO"
    elif TypeOfCase == "No Case" or TypeOfCase == "Non-Litigation Advocacy" or TypeOfCase == "Tenant Rights"  or TypeOfCase == "Rent Strike":
        return "00"
    elif TypeOfCase == "Non-payment":
        return "NP"
    elif TypeOfCase == "Section 8 Administrative Proceeding" or TypeOfCase == "Sec. 8 Termination" or TypeOfCase == "Section 8 Grievance" or TypeOfCase == "Section 8 HQS" or TypeOfCase == "" or TypeOfCase == "Section 8 other" or TypeOfCase == "Section 8 share":
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
    elif TypeOfCase == "HRA Fair Hearing" or TypeOfCase == "Human Rights Complaint" or TypeOfCase == "Mitchell-Lama RFM"or TypeOfCase == "Mitchell-Lama Termination"or TypeOfCase == "NYCHA Housing Grievance"or TypeOfCase == "NYCHA RFM"or TypeOfCase == "Other Administrative Proceeding" or TypeOfCase == "PA Issue: City FEPS/SEPS" or TypeOfCase == "PA Issue: Budgeting" or TypeOfCase == "PA Issue: FEPS" or TypeOfCase == "" or TypeOfCase == "PA Issue: LINC" or TypeOfCase == "PA Issue: Other" or TypeOfCase == "PA Issue: RAU" or TypeOfCase == "PA Issue: Underpayment" or TypeOfCase == "SCRIE/DRIE" or TypeOfCase == "Certificate of No Harassment Case":
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

#Case posture on eligibility date (on trial, no stipulation etc.) - transform them into the HRA initials
def PostureOnEligibility(Posture):
    if Posture == "No Stipulation; No Judgment":
        return "NSNJ"
    elif Posture == "Post-Stipulation, No Judgment":
        return "PSNJ"
    elif Posture == "Post-Judgment, Tenant out of Possession":
        return "PJP"
    elif Posture == "On for Trial":
        return "OFT"
    elif Posture == "Post-Judgment, Tenant in Possession-Judgment Due To Default":
        return "PJD"
    elif Posture == "Post-Judgment, Tenant in Possession-Judgment Due To Other":
        return "PJO"
 
 
#Level of Service becomes Service type - lots of level of service in LS, mapped to Advice Only, Pre-Lit Strategies (brief service, out of court advocacy, hold for review), and Full Rep (mapping to be confirmed)
def ServiceType(LevelOfService):
    if LevelOfService == "Advice":
        return "Advice Only"
    elif LevelOfService == "Brief Service" or LevelOfService == "Out-of-Court Advocacy":
        return "Pre-Litigation Strategies"
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
    elif HousingType == "Public Housing" or HousingType == "Public Housing/NYCHA" or HousingType == "NYCHA/NYCHA" or HousingType == "" or HousingType == "" or HousingType == "" or HousingType == "" or HousingType == "" or HousingType == "" or HousingType == "":
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
    
    elif ReferralSource == "3-1-1" or ReferralSource == "ADP Hotline" or ReferralSource == "Community Organization" or ReferralSource == "Elected Official" or ReferralSource == "Foreclosure" or ReferralSource == "Friends/Family" or ReferralSource == "Home base" or ReferralSource == "In-House" or ReferralSource == "Other City Agency" or ReferralSource == "Outreach" or ReferralSource == "Returning Client" or ReferralSource == "School" or ReferralSource == "Self-referred" or ReferralSource == "Word of mouth" or ReferralSource == "Legal Services":
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
    if ServicesRendered == "Secured Rent Abatement":
        return "Abatement"
    elif ServicesRendered == "Secured Rent Reduction":
        return "Reduction"
    elif ServicesRendered == "Secured Order or Agreement for Repairs in Apartment/Building":
        return "Repairs"
    elif ServicesRendered == "Returned Unit to Rent Regulation":
        return "Regulation"
    elif ServicesRendered == "Obtained Renewal of Lease":
        return "Renewal"
    elif ServicesRendered == "Obtain Ongoing Rent Subsidy":
        return "Subsidy"
    elif ServicesRendered == "Client Security Deposit Returned":
        return "Deposit"
    elif ServicesRendered == "Case Discontinued/Dismissed/Landlord Fails to Prosecute":
        return "Discontinued"
    elif ServicesRendered == "Case Resolved without Judgment of Eviction Against Client":
        return "Resolved"
    elif ServicesRendered == "Secured 6 Months or Longer in Residence":
        return "6Months"
    elif ServicesRendered == "Obtained Succession Rights to Residence":
        return "Succession"
    elif ServicesRendered == "Obtained Negotiated Buyout":
        return "Buyout"
    elif ServicesRendered == "Restored Access to Personal Property":
        return "Property"
    elif ServicesRendered == "Overcame Housing Discrimination":
        return "Discrimination"
    elif ServicesRendered == "Provided Housing-related Consumer Debt Legal Assistance":
        return "Debt"
        
#Mapped to what HRA wants - some of the options are in LegalServer,
def Activities(Activity):
    if Activity == "Counsel Assisted in Filing or Refiling of Answer":
        return "Answer"
    elif Activity == "Filed/Argued/Supplemented Dispositive or other Substantive Motion":
        return "Motion"
    elif Activity == "Filed for an Emergency Order to Show Cause":
        return "OSC"
    elif Activity == "Conducted Traverse Hearing":
        return "Traverse"
    elif Activity == "Conducted Evidentiary Hearing":
        return "Evidentiary"
    elif Activity == "Commenced Trial":
        return "Trial"
    elif Activity == "Filed Appeal":
        return "Appeal"
        
#Subsidy type - if it's not in the HRA list, it has to be 'none' (other is not valid) - they want a smaller list than we record. (mapping to be confirmed)        
def SubsidyType(SubsidyType):
    if SubsidyType == "LINC" or SubsidyType == "HOMETBRA" or SubsidyType == "FEPS" or SubsidyType == "SEPS" or SubsidyType == "City FEPS" or SubsidyType == "HASA" or SubsidyType == "Pathways Home" or SubsidyType == "SOTA" or SubsidyType == "City HRA Subsidy":
        return "HRA Subsidy"
    elif SubsidyType == "HUD VASH":
        return "Section 8"
    elif SubsidyType == "ACS Housing Subsidy":
        return "ACS Subsidy"
    else:
        return "None"

#Does Client have an eligibility date prior to March 1st, 2020?
def PreThreeOne(EligibilityDate):
    if isinstance(EligibilityDate, int) == False:
        return "No"
    elif EligibilityDate < 20200301:
        return "Yes"
    elif EligibilityDate >= 20200301:
        return "No"
        
    