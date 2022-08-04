from flask import request, send_from_directory
from app import app, DataWizardTools, HousingToolBox
import pandas as pd

@app.route("/TRCExternalPreparer", methods=['GET', 'POST'])
def TRCExternalPreparer():
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
        
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        
        
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
        
        #If it is a case in Queens it will have neighborhood - change it to say Queens
        df['city'] = df.apply(lambda x: DataWizardTools.QueensConsolidater(x['City']), axis=1)
        
        #Translation based on HRA Specs            
        df['proceeding'] = df.apply(lambda x: HousingToolBox.TRCProceedingType(x['Housing Type Of Case'],x['Legal Problem Code'],x['Housing Level of Service'],x['DateConstruct']), axis=1)

        #if it's a multi-tenant/group case? change it from saying Yes/no to say "no = individual" or 'yes = Group'
        #Also, if it's an eviction case, post 12/1/21, no response, otherwise make it "needs review"
                
        def ProceedingLevel(GroupCase,TypeOfCase,EvictionProceedings):
            if TypeOfCase in EvictionProceedings:
                return ""
            elif GroupCase == "Yes":
                return "Group"
            elif GroupCase == "No":
                return "Individual"
            else:
                return "Needs Review"
        df['proceeding_level'] = df.apply(lambda x: ProceedingLevel(x['Housing Building Case?'], x['proceeding'], HousingToolBox.evictionproceedings), axis=1)
        
        def StandardizeYears (YIA):
            YIA=str(YIA)
            if YIA == "":
                return ""
            elif ',' in YIA:
                YIA=YIA.replace(',','')
                return int(YIA)
            elif int(YIA) <=-1:
                return .5
            else:
                return int(YIA)
                
        df['years_in_apt'] = df.apply(lambda x:StandardizeYears(x['Housing Years Living In Apartment']),axis=1)
            
        
        #df['Housing Years Living In Apartment']= df['Housing Years Living In Apartment'].replace(',','',regex=True,inplace=True)
        #df['Housing Years Living In Apartment']= df['Housing Years Living In Apartment'].astype(str)
        #df['Housing Years Living In Apartment']= df['Housing Years Living In Apartment'].astype(int)
        
        #For years in apartment, negative 1 or less = 0.5***
        #df['years_in_apt'] = df['Housing Years Living In Apartment'].apply(lambda x: .5 if x <= -1 else x)
        
        
        #Case posture on eligibility date (on trial, no stipulation etc.) - transform them into the HRA initials
        df['posture'] = df.apply(lambda x: HousingToolBox.PostureOnEligibility(x['Housing Posture of Case on Eligibility Date']), axis=1)
        
        
        #Level of Service becomes Service type 
        df['service_type'] = df.apply(lambda x: HousingToolBox.TRCServiceType(x['Housing Level of Service'],x['Legal Problem Code'],x['Primary Funding Code'],x['Referral Source'],x['HRA Release?']), axis=1)
        
        #if below 201, = 'Yes' otherwise 'No'
        df['below_200_FPL'] = df['Percentage of Poverty'].apply(lambda x: "Yes" if x < 200 else "No")
    
        #Subsidy type - if it's not in the HRA list, it has to be 'none' (other is not valid) - they want a smaller list than we record. (mapping to be confirmed)
                
        df['subsidy_type'] = df.apply(lambda x: HousingToolBox.SubsidyType(x['Housing Subsidy Type']), axis=1)
        
        
        #Housing Regulation Type: mapping down - we have way more categories, rent regulated, market rate, or other (mapping to be confirmed). can't be blank

        df['housing_type'] = df.apply(lambda x: HousingToolBox.HousingType(x['Housing Form Of Regulation']), axis=1)

        #Referrals need to be one of their specific categories
            
        df['referral_source'] = df.apply(lambda x: HousingToolBox.ReferralMap(x['Referral Source']), axis = 1)
        
        #Housing Outcomes needs mapping for HRA
        df['outcome'] = df.apply(lambda x: HousingToolBox.Outcome(x['Housing Outcome']), axis=1)
        
        #Outcome related things that need mapping   
        df['services_rendered'] = df.apply(lambda x: HousingToolBox.ServicesRendered(x['Housing Services Rendered to Client']), axis=1)

        #Mapped to what HRA wants - some of the options are in LegalServer,

        df['activities'] = df.apply(lambda x: HousingToolBox.Activities(x['Housing Activity Indicators']), axis=1)
        
        df['LS Waiver Type'] = df['Housing TRC HRA Waiver Categories']
        df['LS Waiver Date'] = df['Housing Date Of Waiver Approval']
        
        #Apply categorical income waiver date to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def CatWaivDate (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, Prim, LT, Type):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return Date
            elif LoS == "Advice Only" and Date != "":
                return Date
            elif LoS == "Advice Only" and Date == "":
                return Date 
            elif Date != "":
                return Date
            elif Pov >= 201 and Date == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "11/28/2016"
                elif Proceeding == "IL":
                        return "09/20/2021"
                else:
                    return Date
            else:
                return Date
                
        df['waiver_approval_date'] = df.apply(lambda x: CatWaivDate(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Date'],x['proceeding'],x['Primary Funding Code'],x['LT_index'],x['Housing TRC HRA Waiver Categories']), axis=1)
        
        
        
        
        #Apply categorical income waiver date to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def CatWaivDateReas (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, Prim, LT, Type):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return str(Date) + "-Original LS response, no FJC indicators"
            elif LoS == "Advice Only" and Date != "":
                return str(Date) + ", Adv Case w Original LS Date"
            elif LoS == "Advice Only" and Date == "":
                return str(Date) + "Adv Case, no date needed"
            elif Date != "":
                return "Has LS waiver date, " + str(Date)
            elif Pov >= 201 and Date == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "11/28/2016, over income case w 2 FJC indicators"
                elif Proceeding == "IL":
                        return "09/20/2021, over income IL case"
                else:
                    return str(Date) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Date) + "Pov<201"
                
        df['Script waiver date Reason'] = df.apply(lambda x: CatWaivDateReas(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Date'],x['proceeding'],x['Primary Funding Code'],x['LT_index'],x['Housing TRC HRA Waiver Categories']), axis=1)
        
        #Apply categorical income waiver type to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def CatWaivType (Pov, EDC, LoS, CaseType, Ref, Type, Proceeding, Prim, LT):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return Type
            elif LoS == "Advice Only" and Type != "":
                return Type
            elif LoS == "Advice Only" and Type == "":
                return Type
            elif Type != "":
                return Type
            elif Pov >= 201 and Type == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "FJC Waiver"
                elif Proceeding == "IL":
                        return "Income Waiver"
                else:
                    return Type
            else:
                return Type
                
        df['waiver'] = df.apply(lambda x: CatWaivType(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Type'],x['proceeding'],x['Primary Funding Code'],x['LT_index']), axis=1)
        
        #Apply categorical income waiver type reasoning to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def CatWaivTypeReas (Pov, EDC, LoS, CaseType, Ref, Proceeding, Prim, LT, Type):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return str(Type) + "-Original LS response, no FJC indicators"
            elif LoS == "Advice Only" and Type != "":
                return str(Type) + ", Adv Case w Original LS waiver type"
            elif LoS == "Advice Only" and Type == "":
                return str(Type) + "Adv Case, no waiver needed"
            elif Type != "":
                return str(Type) + ", waiver type in LS"
            elif Pov >= 201 and Type == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "FJC Waiver, over income case w 2 FJC indicators"
                elif Proceeding == "IL":
                        return "Income Waiver, over income IL case"
                else:
                    return str(Type) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Type) + "Pov<201"
                
        df['Script waiver type Reason'] = df.apply(lambda x: CatWaivTypeReas(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['proceeding'],x['Primary Funding Code'],x['LT_index'],x['LS Waiver Type']), axis=1)
        
        
        
        
        def DeleteChildren (NumChildren):

            if NumChildren == "":
                return 0
            else:
                return NumChildren
        df['num_children'] = df.apply(lambda x: DeleteChildren(x['num_children']), axis=1)
        

        df['DOB'] = df['Date of Birth']
        
        
        
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
        'Script waiver date Reason',
        'LS Waiver Date',
        'waiver',
        'Script waiver type Reason',
        'LS Waiver Type',
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
        'Legal Problem Code',
        #'Non-Housing Case Tester',
        'Primary Funding Code'
        ]]
        
        
        

        """
        Graveyard
                  _(_)_                          wWWWw   _
      @@@@       (_)@(_)   vVVVv     _     @@@@  (___) _(_)_
     @@()@@ wWWWw  (_)\    (___)   _(_)_  @@()@@   Y  (_)@(_)
      @@@@  (___)     `|/    Y    (_)@(_)  @@@@   \|/   (_)\
       /      Y       \|    \|/    /(_)    \|      |/      |
    \ |     \ |/       | / \ | /  \|/       |/    \|      \|/
jgs \\|//   \\|///  \\\|//\\\|/// \|///  \\\|//  \\|//  \\\|// 
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^        
        
        df['Post 12/1/21 Elig Date?'] = df.apply(lambda x: HousingToolBox.PostTwelveOne(x['DateConstruct']), axis=1)
        
        #Apply categorical income waiver date to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def FY22CatWaivDate (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, Prim, LT, Type):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return Date
            elif LoS == "Advice Only" and Date != "":
                return Date
            elif LoS == "Advice Only" and Date == "":
                return Date 
            elif Date != "":
                return Date
            elif Pov >= 201 and Date == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "11/28/2016"
                elif Proceeding == "IL":
                        return "09/20/2021"
                elif EDC < 20220318:
                    if LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT.startswith('LT') == True:
                        return "09/20/2021"
                    elif Ref == "Documented HRA Referral":
                        return "09/20/2021"
                    elif LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT != "":
                        return Date
                    else:
                        return Date
                else:
                    return Date
            else:
                return Date
        
        #Apply categorical income waiver date to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def FY22CatWaivDateReas (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, Prim, LT, Type):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return str(Date) + "-Original LS response, no FJC indicators"
            elif LoS == "Advice Only" and Date != "":
                return str(Date) + ", Adv Case w Original LS Date"
            elif LoS == "Advice Only" and Date == "":
                return str(Date) + "Adv Case, no date needed"
            elif Date != "":
                return "Has LS waiver date, " + str(Date)
            elif Pov >= 201 and Date == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "11/28/2016, over income case w 2 FJC indicators"
                elif Proceeding == "IL":
                        return "09/20/2021, over income IL case"
                elif EDC < 20220318:
                    if LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT.startswith('LT') == True:
                        return "09/20/2021, pre 3/18 Full Rep EvC w Crt#"
                    elif Ref == "Documented HRA Referral":
                        return "09/20/2021, pre 3/18 non Adv Doc HRA Ref"
                    elif LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT != "":
                        return "Full Rep EvC, Check Gen Case Index, may have cat. Income date 09/20/2021"
                    else:
                        return str(Date) + "pre 3/18, may need waiver, not categorical"
                else:
                    return str(Date) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Date) + "Pov<201"
        
        def FY22CatWaivType (Pov, EDC, LoS, CaseType, Ref, Type, Proceeding, Prim, LT):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return Type
            elif LoS == "Advice Only" and Type != "":
                return Type
            elif LoS == "Advice Only" and Type == "":
                return Type
            elif Type != "":
                return Type
            elif Pov >= 201 and Type == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "FJC Waiver"
                elif Proceeding == "IL":
                        return "Income Waiver"
                elif EDC < 20220318:
                    if LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT.startswith('LT') == True:
                        return "Income Waiver"
                    elif Ref == "Documented HRA Referral":
                        return "Income Waiver"
                    elif LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT != "":
                        return Type
                    else:
                        return Type
                else:
                    return Type
            else:
                return Type
        
        def FY22CatWaivTypeReas (Pov, EDC, LoS, CaseType, Ref, Proceeding, Prim, LT, Type):
            EDC = int(EDC)
            LT = str(LT)
            if Ref != "Family Justice Center" and "3011" not in Prim and Type == "FJC Waiver":
                return str(Type) + "-Original LS response, no FJC indicators"
            elif LoS == "Advice Only" and Type != "":
                return str(Type) + ", Adv Case w Original LS waiver type"
            elif LoS == "Advice Only" and Type == "":
                return str(Type) + "Adv Case, no waiver needed"
            elif Type != "":
                return str(Type) + ", waiver type in LS"
            elif Pov >= 201 and Type == "":
                if Ref == "Family Justice Center" and "3011" in Prim:
                    return "FJC Waiver, over income case w 2 FJC indicators"
                elif Proceeding == "IL":
                        return "Income Waiver, over income IL case"
                elif EDC < 20220318:
                    if LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT.startswith('LT') == True:
                        return "Income Waiver, pre 3/18 Full Rep EvC w Crt#"
                    elif Ref == "Documented HRA Referral":
                        return "Income Waiver, pre 3/18 non Adv Doc HRA Ref"
                    elif LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT != "":
                        return "Full Rep EvC, Check Gen Case Index, may have cat. Income Waiver"
                    else:
                        return str(Type) + "pre 3/18, may need waiver, not categorical"
                else:
                    return str(Type) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Type) + "Pov<201"

        #Flag cases that don't have housing-based legal problem codes for Kim's review, after 9/30
        
        def NonHousing (LegalProblemCode):
            if LegalProblemCode.startswith('6') == True:
                return 'Housing'
            else:
                return 'Needs Review'
    
        df['Non-Housing Case Tester'] = df.apply(lambda x: NonHousing(x['Legal Problem Code']), axis=1)
        
        def FY22HousholdSum (PostTwelveOne, ServiceType,  NumAdults, NumChildren, PrimaryFunding):
            if NumChildren == "":
                NumChildren = 0
            if PostTwelveOne == "No":
                if ServiceType == "Advice Only" and PrimaryFunding != "3011 TRC FJC Initiative":
                    return NumAdults + int(NumChildren)
                else:
                    return NumAdults
            else:
                return NumAdults
        df['num_adults'] = df.apply(lambda x: HousholdSum(x['Post 12/1/21 Elig Date?'], x['service_type'], x['num_adults'], x['num_children'],x['Primary Funding Code']), axis=1)
        
         def FY22DeleteChildren (PostTwelveOne, ServiceType, NumChildren, PrimaryFunding):
            if PostTwelveOne == "No":
                if ServiceType == "Advice Only" and PrimaryFunding != "3011 TRC FJC Initiative":
                    return ""
                elif NumChildren == "":
                    return 0
                else:
                    return NumChildren
            elif NumChildren == "":
                return 0
            else:
                return NumChildren
        df['num_children'] = df.apply(lambda x: DeleteChildren(x['Post 12/1/21 Elig Date?'], x['service_type'], x['num_children'],x['Primary Funding Code']), axis=1)
        
        #Only have to report birth year 
        def FY22RedactBirthday(PostTwelveOne,ServiceType,DOB,PrimaryFunding):
            if PostTwelveOne == "No":
                if ServiceType == "Advice Only" and PrimaryFunding != "3011 TRC FJC Initiative":
                    return DOB[6:]
                else:
                    return DOB
            else:
                return DOB
        df['DOB'] = df.apply(lambda x: RedactBirthday(x['Post 12/1/21 Elig Date?'], x['service_type'], x['Date of Birth'],x['Primary Funding Code']), axis=1)
        
        
        def RedactLT(PostTwelveOne, ServiceType, ToRedact, PrimaryFunding, ProblemCode, Consent):
            if PostTwelveOne == "No":
                if ServiceType == 'Advice Only':
                    if Consent == "Yes":
                        return ToRedact
                    else:
                        return ""
                else:
                    return ToRedact
            else:
                return ToRedact
        
        df['LT_index'] = df.apply(lambda x: RedactLT(x['Post 12/1/21 Elig Date?'], x['service_type'], x['LT_index'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?']), axis=1)
        
        #Redacting Function Blank
        def RedactAnything(PostTwelveOne, ServiceType, ToRedact, PrimaryFunding, ProblemCode, Consent, Referral):
            if PostTwelveOne == "No":
                if ServiceType == 'Advice Only':
                    if PrimaryFunding == "3011 TRC FJC Initiative" or Referral.startswith("FJC") == True:
                        if Consent == "Yes":
                            return ToRedact
                        else:
                            return ""
                    else:
                        return ""
                else:
                    return ToRedact
            else:
                return ToRedact
        df['DHCI'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['Housing Signed DHCI Form'], x['Primary Funding Code'], x['Legal Problem Code'], x['HRA Release?'], x['Referral Source']), axis=1)
        
        #No names, (not full date etc.)
        df['first_name'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['Client First Name'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        df['last_name'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['Client Last Name'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        
        #also redact PA#, SS#, LT#, address, monthly rent, individual or group, years in apt, referral source, annual income, DHCI, posture of case on eligibility, at or below 200%, # of units in buildling, subsidy type, housing type, outcome, outcome date, services renderd to client, activity indicators, 
        
        df['PA_number'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['PA_number'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        
        df['SSN'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['SSN'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        
        df['Street'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['Street'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
         
        df['Unit'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['Unit'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
          
        #df['city'] = df.apply(lambda x: RedactAnything(x['service_type'], x['city'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?']), axis=1)
           
        df['street_number'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['street_number'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
            
        df['rent'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['rent'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        
       
         
        df['proceeding_level'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['proceeding_level'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
          
        df['years_in_apt'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['years_in_apt'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
           
        df['referral_source'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['referral_source'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
            
        df['income'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['income'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
             
        df['DHCI'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['DHCI'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        
        df['posture'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['posture'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
         
        #df['below_200_FPL'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['below_200_FPL'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?']), axis=1)
          
        df['units_in_bldg'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['units_in_bldg'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        
        df['subsidy_type'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['subsidy_type'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
         
        df['housing_type'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['housing_type'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
          
        df['outcome_date'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['outcome_date'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
           
        df['outcome'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['outcome'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        
        df['services_rendered'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['services_rendered'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
           
        df['activities'] = df.apply(lambda x: RedactAnything(x['Post 12/1/21 Elig Date?'], x['service_type'], x['activities'], x['Primary Funding Code'],x['Legal Problem Code'],x['HRA Release?'], x['Referral Source']), axis=1)
        """
      
        
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
        #(first_row, first_col, last_row, last_col)
        shape = (df.shape[0]+1)
        ColShape= int(df.shape[1])
        HLLoc=df.columns.get_loc("Hyperlinked CaseID#")
        #worksheet.set_column ('AN:AN',30,link_format)
        worksheet.set_column (HLLoc,HLLoc,30,link_format)
        worksheet.conditional_format('C2:BO100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
        worksheet.conditional_format(0,0,0,ColShape,{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'LS',
                                                 'format': problem_format})
        worksheet.conditional_format(0,0,0,ColShape,{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Reason',
                                                 'format': problem_format})
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Formatted " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>TRC Report Preparer</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Prepare Cases for TRC External Report:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=TRC-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1969" target="_blank">TRC External Report</a>.</li>
    <li>This tool can also be used with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1507" target="_blank">TRC Raw Case Data Report</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    