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