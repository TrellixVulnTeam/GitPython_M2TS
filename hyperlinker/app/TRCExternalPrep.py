from flask import request, send_from_directory
from app import app, DataWizardTools, HousingTools
import pandas as pd

@app.route("/TRCExternalPrep", methods=['GET', 'POST'])
def TRCExternalPrep():
    #upload file from computer via browser
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        #Import Excel Sheet into dataframe
        df = pd.read_excel(f,skiprows=2)
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        #Create Hyperlinks
        df['Hyperlinked CaseID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)          

        ###This is where all the functions happen:###
        
        #Just direct mapping for new column names
        df['first_name'] = df['Client First Name']
        df['last_name'] = df['Client Last Name']
        df['SSN'] = df['Social Security #']
        df['PA_number'] = df['Gen Pub Assist Case Number']
        df['DOB'] = df['Date of Birth']
        df['num_adults'] = df['Number of People 18 and Over']
        df['num_children'] = df['Number of People under 18']
        df['Unit'] = df['Apt#/Suite#']
        df['city'] = df['City']
        df['zip'] = df['Zip Code']
        df['waiver_approval_date'] = df['Housing Date Of Waiver Approval']
        df['waiver'] = df['Housing TRC HRA Waiver Categories']
        df['rent'] = df['Housing Total Monthly Rent']
        df['LT_index'] = df['Gen Case Index Number']
        df['language'] = df['Language']
        df['income'] = df['Total Annual Income ']
        df['eligibility_date'] = df['HAL Eligibility Date']
        df['DHCI'] = df['Housing Signed DHCI Form']
        df['units_in_bldg'] = df['Housing Number Of Units In Building']
        df['outcome_date'] = df['Housing Outcome Date']
        
        #Append the 'LSNYC' prefix to the caseIDs we submit
        df['id'] = 'LSNYC' + df['Matter/Case ID#']
        
        #Turn our funding codes into HRA Program Names
        #*for trc (3018 and 3011) everything is AHTP - more complicated for UA etc.
        df['program_name'] = 'AHTP'
        
        #Separate out street number from street name (based on first space)
        df['street_number'] = df['Street Address'].str.split(' ').str[0]
        df['Street'] = df['Street Address'].str.split(' ',1).str[1]
        
        #Translation based on HRA Specs            
        df['proceeding'] = df.apply(lambda x: HousingTools.ProceedingType(x['Housing Type Of Case']), axis=1)

        #if it's a multi-tenant/group case? change it from saying Yes/no to say "no = individual" or 'yes = Group'
        #Also, if it's an eviction case, it's individual, otherwise make it "needs review"
        
        def ProceedingLevel(GroupCase,TypeOfCase,EvictionProceedings):
            if GroupCase == "Yes":
                return "Group"
            elif GroupCase == "No":
                return "Individual"
            elif TypeOfCase in EvictionProceedings:
                return "Individual"
            else:
                return "Needs Review"
        df['proceeding_level'] = df.apply(lambda x: ProceedingLevel(x['Housing Building Case?'], x['proceeding'], HousingTools.evictionproceedings), axis=1)
        
        #For years in apartment, negative 1 or less = 0.5
        df['years_in_apt'] = df['Housing Years Living In Apartment'].apply(lambda x: .5 if x <= -1 else x)
        
        
        #Case posture on eligibility date (on trial, no stipulation etc.) - transform them into the HRA initials
        df['posture'] = df.apply(lambda x: HousingTools.PostureOnEligibility(x['Housing Posture of Case on Eligibility Date']), axis=1)
        
        
        #Level of Service becomes Service type 
        df['service_type'] = df.apply(lambda x: HousingTools.ServiceType(x['Housing Level of Service']), axis=1)
        
        #if below 201, = 'Yes' otherwise 'No'
        df['below_200_FPL'] = df['Percentage of Poverty'].apply(lambda x: "Yes" if x < 200 else "No")
    
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
                
        df['subsidy_type'] = df.apply(lambda x: SubsidyType(x['Housing Subsidy Type']), axis=1)
        
        
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
        
        df['housing_type'] = df.apply(lambda x: HousingType(x['Housing Form Of Regulation']), axis=1)

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
            
        df['referral_source'] = df.apply(lambda x: ReferralMap(x['Referral Source']), axis = 1)
        
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
        df['outcome'] = df.apply(lambda x: Outcome(x['Housing Outcome']), axis=1)


        
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

        df['services_rendered'] = df.apply(lambda x: ServicesRendered(x['Housing Services Rendered to Client']), axis=1)

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
        
        df['activities'] = df.apply(lambda x: Activities(x['Housing Activity Indicators']), axis=1)
        
        
        #Differentiate pre- and post- 3/1/20 eligibility date cases
        
           
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        
        df['Pre-3/1/20 Elig Date?'] = df['DateConstruct'].apply(lambda x: "Yes" if x <20200301 else "No")
        
        
        #different guidelines for post 3/1/20 eligibility dates
        #mostly for advice cases for TRC (for UA also for out of court & brief)

        #No names, Only have to report birth year (not full date etc.) - or just #give them whole date without name

        #Total household goes into adult column
        
        
        
        
        

        ###Finalizing Report###
        #put columns in correct order
        
        df = df[['id',
        'program_name',
        'first_name',
        'last_name',
        'SSN',
        'PA_number',
        'DOB',
        'num_adults',
        'num_children',
        'street_number',
        'Street',
        'Unit',
        'city',
        'zip',
        'waiver_approval_date',
        'waiver',
        'rent',
        'proceeding',
        'LT_index',
        'proceeding_level',
        'years_in_apt',
        'language',
        'referral_source',
        'income',
        'eligibility_date',
        'DHCI',
        'posture',
        'service_type',
        'below_200_FPL',
        'units_in_bldg',
        'subsidy_type',
        'housing_type',
        'outcome_date',
        'outcome',
        'services_rendered',
        'activities',
        'HRA Release?',
        'Percentage of Poverty',
        'Primary Advocate',
        'Hyperlinked CaseID#',
        'Pre-3/1/20 Elig Date?'
        ]]
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        #highlight yellow if needs review
        #make columns wider
        #give the hyperlink format
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        problem_format = workbook.add_format({'bg_color':'yellow'})
        worksheet.freeze_panes(1,0)
        worksheet.set_column('A:BL',20)
        worksheet.set_column ('AN:AN',30,link_format)
        worksheet.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Formatted " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>TRC Report Prep</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Prep Cases for TRC External Report:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=TRC-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1969" target="_blank">TRC External Report</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
