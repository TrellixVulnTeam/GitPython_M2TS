from flask import request, send_from_directory
from app import app, DataWizardTools, HousingTools
import pandas as pd

@app.route("/TRCExternalPrepCovid", methods=['GET', 'POST'])
def TRCExternalPrepCovid():
    #upload file from computer via browser
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        
        #turn the excel file into a dataframe, but skip the top 2 rows if they are blank
        test = pd.read_excel(f)
        test.fillna('',inplace=True)
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
                
        df['subsidy_type'] = df.apply(lambda x: HousingTools.SubsidyType(x['Housing Subsidy Type']), axis=1)
        
        
        #Housing Regulation Type: mapping down - we have way more categories, rent regulated, market rate, or other (mapping to be confirmed). can't be blank

        df['housing_type'] = df.apply(lambda x: HousingTools.HousingType(x['Housing Form Of Regulation']), axis=1)

        #Referrals need to be one of their specific categories
            
        df['referral_source'] = df.apply(lambda x: HousingTools.ReferralMap(x['Referral Source']), axis = 1)
        
        #Housing Outcomes needs mapping for HRA
        df['outcome'] = df.apply(lambda x: HousingTools.Outcome(x['Housing Outcome']), axis=1)
        
        #Outcome related things that need mapping   
        df['services_rendered'] = df.apply(lambda x: HousingTools.ServicesRendered(x['Housing Services Rendered to Client']), axis=1)

        #Mapped to what HRA wants - some of the options are in LegalServer,

        df['activities'] = df.apply(lambda x: HousingTools.Activities(x['Housing Activity Indicators']), axis=1)
        
        
        #Differentiate pre- and post- 3/1/20 eligibility date cases
           
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        df['Pre-3/1/20 Elig Date?'] = df.apply(lambda x: HousingTools.PreThreeOne(x['DateConstruct']), axis=1)
        
        
        
        ##different guidelines for post 3/1/20 eligibility dates
        ##If case is advice and has a post-3/1 eligibility date
        
        #Sum household in adult column and leave children blank
        def HousholdSum (ServiceType, PreThreeOne, NumAdults, NumChildren):
            if ServiceType == "Advice Only" and PreThreeOne == "No":
                return NumAdults + NumChildren
            else:
                return NumAdults
        df['num_adults'] = df.apply(lambda x: HousholdSum(x['service_type'], x['Pre-3/1/20 Elig Date?'], x['num_adults'], x['num_children']), axis=1)
        
        def DeleteChildren (ServiceType, PreThreeOne, NumChildren):
            if ServiceType == "Advice Only" and PreThreeOne == "No":
                return ""
            else:
                return NumChildren
        df['num_children'] = df.apply(lambda x: DeleteChildren(x['service_type'], x['Pre-3/1/20 Elig Date?'], x['num_children']), axis=1)
        
        #Only have to report birth year 
        def RedactBirthday(ServiceType, PreThreeOne,DOB):
            if ServiceType == "Advice Only" and PreThreeOne == "No":
                return "01/01/"+ DOB[6:]
            else:
                return DOB
        df['DOB'] = df.apply(lambda x: RedactBirthday(x['service_type'], x['Pre-3/1/20 Elig Date?'], x['Date of Birth']), axis=1)
        
        #DHCI Blank
        def RedactAnything(ServiceType, PreThreeOne, ToRedact):
            if ServiceType == "Advice Only" and PreThreeOne == "No":
                return ""
            else:
                return ToRedact
        df['DHCI'] = df.apply(lambda x: RedactAnything(x['service_type'], x['Pre-3/1/20 Elig Date?'], x['Housing Signed DHCI Form']), axis=1)
        
        #No names, (not full date etc.) - or just #give them whole date without name
        df['first_name'] = df.apply(lambda x: RedactAnything(x['service_type'], x['Pre-3/1/20 Elig Date?'], x['Client First Name']), axis=1)
        df['last_name'] = df.apply(lambda x: RedactAnything(x['service_type'], x['Pre-3/1/20 Elig Date?'], x['Client Last Name']), axis=1)
        
        
        

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
    <title>TRC Report Prep [COVID]</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Prep Cases for TRC External Report [COVID]:</h1>
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
    
