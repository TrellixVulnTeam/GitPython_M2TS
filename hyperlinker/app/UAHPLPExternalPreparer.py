from flask import request, send_from_directory
from app import app, DataWizardTools, HousingToolBox
import pandas as pd
import numpy as np

@app.route("/UAHPLPExternalPreparer", methods=['GET', 'POST'])
def UAHPLPExternalPreparer():
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
        
        df['zip'] = df['Zip Code'].apply(lambda x: '{0:0>5}'.format(x))
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
        
        #Translation based on HRA Specs  (this got moved up cuz it's output is necessary for program name)           
        df['proceeding'] = df.apply(lambda x: HousingToolBox.UACProceedingType(x['Housing Type Of Case'],x['Legal Problem Code'],x['Close Reason'],x['Housing Level of Service']), axis=1)
        
        #cases in certain zip codes (RTC zips) that are eviction are UA - everything else is non-UA **Bounce this to housing tools**
        #***should this be changed for nonhousing and/or bounced to housing tools?
        
        def UAorNonUA (TypeOfCase):

            if TypeOfCase in HousingToolBox.evictionproceedings:
                return "UA"
            else:
                return "Non-UA"

        df['program_name'] = df.apply(lambda x: UAorNonUA(x['proceeding']), axis=1)
        
        
        
        #Separate out street number from street name (based on first space)
        df['street_number'] = df['Street Address'].str.split(' ').str[0]
        df['Street'] = df['Street Address'].str.split(' ',1).str[1]
        
        #If it is a case in Queens it will have neighborhood - change it to say Queens
        df['city'] = df.apply(lambda x: DataWizardTools.QueensConsolidater(x['City']), axis=1)

        #if it's a multi-tenant/group case? change it from saying Yes/no to say "no = individual" or 'yes = Group'
        #Also, if it's an eviction case, it's individual, otherwise make it "needs review"
        
        def ProceedingLevel(GroupCase,TypeOfCase,EvictionProceedings):
            if TypeOfCase in EvictionProceedings:
                return ""
            elif GroupCase == "Yes":
                return "Group"
            elif GroupCase == "No":
                return "Individual"
            else:
                return ""
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
        
        
        #For years in apartment, negative 1 or less = 0.5
        #df['years_in_apt'] = df['Housing Years Living In Apartment'].apply(lambda x: .5 if x <= -1 else x)
        
        
        #Case posture on eligibility date (on trial, no stipulation etc.) - transform them into the HRA initials
        df['posture'] = df.apply(lambda x: HousingToolBox.PostureOnEligibility(x['Housing Posture of Case on Eligibility Date']), axis=1)
        
        
        #Level of Service becomes Service type 
        df['service_type'] = df.apply(lambda x: HousingToolBox.UACServiceType(x['Housing Level of Service'],x['program_name'],x['Close Reason'],x['Legal Problem Code']), axis=1)
        
        
        df['Percentage of Poverty'] = df['Percentage of Poverty'].apply(lambda x: 0 if x == '' else x)
        #if below 201, = 'Yes' otherwise 'No'
        df['below_200_FPL'] = df['Percentage of Poverty'].apply(lambda x: "Yes" if x < 200 else "No")
    
        #Subsidy type - if it's not in the HRA list, it has to be 'none' (other is not valid) - they want a smaller list than we record. (mapping to be confirmed)
                
        df['subsidy_type'] = df.apply(lambda x: HousingToolBox.SubsidyType(x['Housing Subsidy Type']), axis=1)
        
        
        #Housing Regulation Type: mapping down - we have way more categories, rent regulated, market rate, or other (mapping to be confirmed). can't be blank

        df['housing_type'] = df.apply(lambda x: HousingToolBox.HousingType(x['Housing Form Of Regulation']), axis=1)

        #Referrals need to be one of their specific categories
            
        df['referral_source'] = df.apply(lambda x: HousingToolBox.ReferralMap(x['Referral Source']), axis = 1)
        
        #Differentiate pre- and post- 3/1/20 12/1/21 eligibility date cases          
        df['DateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        
                
        #Waiver Section
        df['Housing Date Of Waiver Approval'] = df['Housing Date Of Waiver Approval'].replace(pd.NaT, '')
        
        def DateCheck (EligibilityDate):
            try:
                EligibilityDate = EligibilityDate.strftime('%m/%d/%Y')
                return EligibilityDate
                #print("Converted dates to strings")
            except:
                #print("Did not convert dates to strings")
                return EligibilityDate
                
        df['Housing Date Of Waiver Approval'] = df.apply(lambda x: DateCheck(x['Housing Date Of Waiver Approval']), axis=1)
        
        df['LS Waiver Type'] = df['Housing TRC HRA Waiver Categories']
        df['LS Waiver Date'] = df['Housing Date Of Waiver Approval']
        df['LS Waiver Date'] = df['LS Waiver Date'].replace(pd.NaT, '')
        #df['LS Waiver Date'].replace("NaT",np.NaN, inplace=True)
        #df.fillna("", inplace=True)
        #print(df['LS Waiver Date'])
        
        #Apply categorical income waiver date to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def CatWaivDate (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Date != "":
                return Date
            elif LoS == "Advice Only" and Date == "":
                return Date
            elif LoS == "Brief Legal Assistance" and Date != "":
                return Date
            elif LoS == "Brief Legal Assistance" and Date == "":
                return Date
            elif Date != "":
                return Date
            elif Pov >= 201 and Date == "":
                if Proceeding == "IL":
                        return "09/20/2021"
                else:
                    return Date
            else:
                return Date
                
        df['waiver_approval_date'] = df.apply(lambda x: CatWaivDate(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Date'],x['proceeding'],x['LT_index']), axis=1)
                
        def CatWaivDateReas (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Date != "":
                return str(Date) + ", Adv Case w Original LS Date"
            elif LoS == "Advice Only" and Date == "":
                return str(Date) + "Adv Case, no date needed"
            elif LoS == "Brief Legal Assistance" and Date != "":
                return str(Date) + ", Brief Case w Original LS Date"
            elif LoS == "Brief Legal Assistance" and Date == "":
                return str(Date) + "Brief Case, no date needed"
            elif Date != "":
                return "Has LS waiver date, " + str(Date)
            elif Pov >= 201 and Date == "":
                if Proceeding == "IL":
                    return "09/20/2021, over income IL case"
                else:
                    return str(Date) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Date) + "Pov<201"
                
        df['Script waiver date Reason'] = df.apply(lambda x: CatWaivDateReas(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Date'],x['proceeding'],x['LT_index']), axis=1)
        
        def CatWaivTypeReas (Pov, EDC, LoS, CaseType, Ref, Type, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Type != "":
                return str(Type) + ", Adv Case w Original LS Waiver"
            elif LoS == "Advice Only" and Type == "":
                return str(Type) + "Adv Case, no type needed"
            elif LoS == "Brief Legal Assistance" and Type != "":
                return str(Type) + ", Brief Case w Original LS Waiver"
            elif LoS == "Brief Legal Assistance" and Type == "":
                return str(Type) + "Brief Case, no type needed"
            elif Type != "":
                return "Has LS waiver type, " + str(Type)
            elif Pov >= 201 and Type == "":
                if Proceeding == "IL":
                    return "Income Waiver, over income IL case"
                else:
                    return str(Type) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Type) + "Pov<201"
                
        df['Script waiver type Reason'] = df.apply(lambda x: CatWaivTypeReas(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Type'],x['proceeding'],x['LT_index']), axis=1)
        
        #Apply categorical income waiver date to cases that had categorical waivers 9/20/21
        def CatWaivType (Pov, EDC, LoS, CaseType, Ref, Type, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Type != "":
                return Type
            elif LoS == "Advice Only" and Type == "":
                return Type
            elif LoS == "Brief Legal Assistance" and Type != "":
                return Type
            elif LoS == "Brief Legal Assistance" and Type == "":
                return Type
            elif Type != "":
                return Type
            elif Pov >= 201 and Type == "":
                if Proceeding == "IL":
                    return "Income Waiver"
                else:
                    return Type
            else:
                return Type

        df['waiver'] = df.apply(lambda x: CatWaivType(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Type'],x['proceeding'],x['LT_index']), axis=1)        
        
        
                
        #Housing Outcomes needs mapping for HRA
        df['outcome'] = df.apply(lambda x: HousingToolBox.Outcome(x['Housing Outcome']), axis=1)
        
        #Outcome related things that need mapping   
        df['services_rendered'] = df.apply(lambda x: HousingToolBox.ServicesRendered(x['Housing Services Rendered to Client']), axis=1)

        #Mapped to what HRA wants - some of the options are in LegalServer,

        df['activities'] = df.apply(lambda x: HousingToolBox.Activities(x['Housing Activity Indicators']), axis=1)
        

        
        #Map 'Borough Values' based on Zip Code
        
        df['BoroughByZip'] = df.apply(lambda x: DataWizardTools.ZipToCity(x['Zip Code']), axis=1)
        
        ##different guidelines for post 3/1/20 eligibility dates
        ##If case is advice and has a post-3/1 eligibility date
        
        
        
        def DeleteChildren (NumChildren):
            if NumChildren == "":
                return 0
            else:
                return NumChildren
        df['num_children'] = df.apply(lambda x: DeleteChildren(x['num_children']), axis=1)
        


        df['DOB'] = df['Date of Birth']
        
        
        
        #HRA is tracking things differently than we are - extra column at end so we can see what counts toward what what from their perspective
        
        def NewProgramAssignment(Proceeding):
            if Proceeding == "NP" or Proceeding == "HO" or Proceeding == "IL" or Proceeding == "TT":
                return "UA"
            elif Proceeding == "EA" or Proceeding == "EJ" or Proceeding == "HP" or Proceeding == "DA" or Proceeding == "7A" or Proceeding == "78" or Proceeding == "S8" or Proceeding == "FC" or Proceeding == "OA" or Proceeding == "OS" or Proceeding == "OO":
                return "Non-UA"
        
        df['2020NewProgramAssignment'] = df.apply(lambda x: NewProgramAssignment(x['proceeding']), axis = 1)
        
        #remove BEN cases
        df = df[df['proceeding'] != 'BEN']
        
        
        #remove non pre-12/1 advice case where HRA Consent is not Yes
        #Removes Pre 12/1 non Advice/Brief cases w HRA Blank/No AND post 12/1 HRA Blank/No cases

        def ReleaseRemove(ServiceType, HRARelease):
            if HRARelease == "Yes":
                return "Keep"
            else:
                return "Remove"
        df['Remove post 12-1 Cases?'] = df.apply(lambda x: ReleaseRemove(x['service_type'],x['HRA Release?']), axis=1)
        
        df = df[df['Remove post 12-1 Cases?'] != 'Remove']
        
        #Removes HRA No/Blank Advice/Brief cases w real Gen Case numbers
        def ReleaseRemoveTwo(ServiceType, CaseNum,Release):
            if ServiceType == "Advice Only" or ServiceType == "Brief Service":
                CaseNum = str(CaseNum)
                if CaseNum != "" and CaseNum.startswith("n") == False and CaseNum.startswith("N") == False:
                    if Release == "" or Release == "No":
                        return "Remove"
        df['Remove no release yes index cases'] = df.apply(lambda x: ReleaseRemoveTwo(x['service_type'],x['LT_index'],x['HRA Release?']), axis = 1)
       
        df = df[df['Remove no release yes index cases'] != 'Remove']
        
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
        'Hyperlinked CaseID#',
        #'Pre-3/1/20 Elig Date?',
        #'Post 12/1/21 Elig Date?',
        '2020NewProgramAssignment',
        #'Remove Cases?',
        'Legal Problem Code',
        'BoroughByZip',
        'Assigned Branch/CC',
        #'Remove no release yes index cases'
        ]]
        
        '''
        Graveyard
        
                  _(_)_                          wWWWw   _
      @@@@       (_)@(_)   vVVVv     _     @@@@  (___) _(_)_
     @@()@@ wWWWw  (_)\    (___)   _(_)_  @@()@@   Y  (_)@(_)
      @@@@  (___)     `|/    Y    (_)@(_)  @@@@   \|/   (_)\
       /      Y       \|    \|/    /(_)    \|      |/      |
    \ |     \ |/       | / \ | /  \|/       |/    \|      \|/
    \\|//   \\|///  \\\|//\\\|/// \|///  \\\|//  \\|//  \\\|// 
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        
        #df['Pre-3/1/20 Elig Date?'] = df.apply(lambda x: HousingToolBox.PreThreeOne(x['DateConstruct']), axis=1)
        df['Post 12/1/21 Elig Date?'] = df.apply(lambda x: HousingToolBox.PostTwelveOne(x['DateConstruct']), axis=1)
        
        #old waiver date
        
        EDC = int(EDC)
            LT = str(LT)
            if Pov >= 201 and Date == "":
                if Proceeding == "IL":
                        return "09/20/2021"
                elif CaseType != "Brief Legal Assistance" and CaseType != "Advice Only": 
                    if EDC < 20220318:
                        if LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT.startswith("LT") == True:
                            return "09/20/2021"
                        elif Ref == "Documented HRA Referral":
                            return "09/20/2021"
                        else:
                            return Date
                    else:
                        return Date
                else:
                    return Date
            else:
                return Date
        
        #Apply categorical income waiver date to cases that had categorical waivers dated 9/20/21 (Fy22 to 3/18/22)
        def FY22CatWaivDate (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Date != "":
                return Date
            elif LoS == "Advice Only" and Date == "":
                return Date
            elif LoS == "Brief Legal Assistance" and Date != "":
                return Date
            elif LoS == "Brief Legal Assistance" and Date == "":
                return Date
            elif Date != "":
                return Date
            elif Pov >= 201 and Date == "":
                if Proceeding == "IL":
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
                
        df['waiver_approval_date'] = df.apply(lambda x: CatWaivDate(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Date'],x['proceeding'],x['LT_index']), axis=1)
                
        def FY22CatWaivDateReas (Pov, EDC, LoS, CaseType, Ref, Date, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Date != "":
                return str(Date) + ", Adv Case w Original LS Date"
            elif LoS == "Advice Only" and Date == "":
                return str(Date) + "Adv Case, no date needed"
            elif LoS == "Brief Legal Assistance" and Date != "":
                return str(Date) + ", Brief Case w Original LS Date"
            elif LoS == "Brief Legal Assistance" and Date == "":
                return str(Date) + "Brief Case, no date needed"
            elif Date != "":
                return "Has LS waiver date, " + str(Date)
            elif Pov >= 201 and Date == "":
                if Proceeding == "IL":
                        return "09/20/2021, over income IL case"
                elif EDC < 20220318:
                    if LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT.startswith('LT') == True:
                        return "09/20/2021, pre 3/18 Full Rep EvC w Crt#"
                    elif Ref == "Documented HRA Referral":
                        return "09/20/2021, pre 3/18 non Adv/Brf Doc HRA Ref"
                    elif LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT != "":
                        return "Full Rep EvC, Check Gen Case Index, may have cat. waiver of 09/20/2021"
                    else:
                        return str(Date) + "pre 3/18, may need waiver, not categorical"
                else:
                    return str(Date) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Date) + "Pov<201"
                
        df['Script waiver date Reason'] = df.apply(lambda x: CatWaivDateReas(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Date'],x['proceeding'],x['LT_index']), axis=1)
        
        def FY22CatWaivTypeReas (Pov, EDC, LoS, CaseType, Ref, Type, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Type != "":
                return str(Type) + ", Adv Case w Original LS Waiver"
            elif LoS == "Advice Only" and Type == "":
                return str(Type) + "Adv Case, no type needed"
            elif LoS == "Brief Legal Assistance" and Type != "":
                return str(Type) + ", Brief Case w Original LS Waiver"
            elif LoS == "Brief Legal Assistance" and Type == "":
                return str(Type) + "Brief Case, no type needed"
            elif Type != "":
                return "Has LS waiver type, " + str(Type)
            elif Pov >= 201 and Type == "":
                if Proceeding == "IL":
                        return "Income Waiver, over income IL case"
                elif EDC < 20220318:
                    if LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT.startswith('LT') == True:
                        return "Income Waiver, pre 3/18 Full Rep EvC w Crt#"
                    elif Ref == "Documented HRA Referral":
                        return "Income Waiver, pre 3/18 non Adv/Brf Doc HRA Ref"
                    elif LoS == "Full Rep" and CaseType in HousingToolBox.evictiontypes and LT != "":
                        return "Full Rep EvC, Check Gen Case Index, may have cat. Income Waiver"
                    else:
                        return str(Type) + "pre 3/18, may need waiver, not categorical"
                else:
                    return str(Type) + "post 3/18 case, may need waiver, not categorical"
            else:
                return str(Type) + "Pov<201"
                
        df['Script waiver type Reason'] = df.apply(lambda x: CatWaivTypeReas(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Type'],x['proceeding'],x['LT_index']), axis=1)
        
        #Apply categorical income waiver date to cases that had categorical waivers 9/20/21
        def FY22CatWaivType (Pov, EDC, LoS, CaseType, Ref, Type, Proceeding, LT):
            EDC = int(EDC)
            LT = str(LT)
            if LoS == "Advice Only" and Type != "":
                return Type
            elif LoS == "Advice Only" and Type == "":
                return Type
            elif LoS == "Brief Legal Assistance" and Type != "":
                return Type
            elif LoS == "Brief Legal Assistance" and Type == "":
                return Type
            elif Type != "":
                return Type
            elif Pov >= 201 and Type == "":
                if Proceeding == "IL":
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

        df['waiver'] = df.apply(lambda x: CatWaivType(x['Percentage of Poverty'],x['DateConstruct'],x['service_type'],x['Housing Type Of Case'],x['referral_source'],x['LS Waiver Type'],x['proceeding'],x['LT_index']), axis=1)        
        
        #old CatWaivType script
         EDC = int(EDC)
            LT = str(LT)
            if Pov >= 201 and Type == "":
                if Proceeding == "IL":
                        return "Income Waiver"
                elif CaseType != "Brief Legal Assistance" and CaseType != "Advice Only": 
                    if EDC < 20220318:
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
            else:
                return Type
        
        #Sum household in adult column and leave children blank
        def FY22HousholdSum (ServiceType, PostTwelveOne, NumAdults, NumChildren):
            if NumChildren == "":
                NumChildren = 0
            if ServiceType == "Advice Only" and PostTwelveOne == "No":
                return NumAdults + NumChildren
            elif ServiceType == "Brief Legal Assistance" and PostTwelveOne == "No":
                return NumAdults + NumChildren
            else:
                return NumAdults
        df['num_adults'] = df.apply(lambda x: HousholdSum(x['service_type'], x['Post 12/1/21 Elig Date?'], x['num_adults'], x['num_children']), axis=1)
        
        def FY22DeleteChildren (ServiceType, PostTwelveOne, NumChildren):
            if ServiceType == "Advice Only" and PostTwelveOne == "No":
                return ""
            elif ServiceType == "Brief Legal Assistance" and PostTwelveOne == "No":
                return ""
            elif NumChildren == "":
                return 0
            else:
                return NumChildren
        df['num_children'] = df.apply(lambda x: DeleteChildren(x['service_type'], x['Post 12/1/21 Elig Date?'], x['num_children']), axis=1)
        
        #Only have to report birth year 
        def FY22RedactBirthday(ServiceType, PostTwelveOne,DOB):
            if ServiceType == "Advice Only" and PostTwelveOne == "No":
                return DOB[6:]
            if ServiceType == "Brief Legal Assistance" and PostTwelveOne == "No":
                return DOB[6:]
            else:
                return DOB
        df['DOB'] = df.apply(lambda x: RedactBirthday(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Date of Birth']), axis=1)
        
        
        #DHCI Blank
        def RedactAnything(ServiceType, PostTwelveOne, ToRedact):
            if ServiceType == "Advice Only" and PostTwelveOne == "No":
                return ""
            elif ServiceType == "Brief Legal Assistance" and PostTwelveOne == "No":
                return ""
            else:
                return ToRedact
        df['DHCI'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Housing Signed DHCI Form']), axis=1)
        
        #No names, (not full date etc.)
        df['first_name'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Client First Name']), axis=1)
        df['last_name'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Client Last Name']), axis=1)
        
        #also redact PA#, SS#, LT#, address, monthly rent, individual or group, years in apt, referral source, annual income, DHCI, posture of case on eligibility, at or below 200%, # of units in buildling, subsidy type, housing type, outcome, outcome date, services renderd to client, activity indicators, 
        
        df['PA_number'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['PA_number']), axis=1)
        
        df['SSN'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['SSN']), axis=1)
        
        df['Street'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Street']), axis=1)
         
        df['Unit'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['Unit']), axis=1)
          
        #df['city'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['city']), axis=1)
           
        df['street_number'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['street_number']), axis=1)
            
        df['rent'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['rent']), axis=1)
        
        df['LT_index'] = df.apply(lambda x: HousingToolBox.NoReleaseRedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['LT_index'], x['HRA Release?']), axis=1)
         
        df['proceeding_level'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['proceeding_level']), axis=1)
          
        df['years_in_apt'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['years_in_apt']), axis=1)
           
        df['referral_source'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['referral_source']), axis=1)
            
        df['income'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['income']), axis=1)
             
        df['DHCI'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['DHCI']), axis=1)
        
        df['posture'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['posture']), axis=1)
         
        #df['below_200_FPL'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['below_200_FPL']), axis=1)
          
        df['units_in_bldg'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['units_in_bldg']), axis=1)
        
        df['subsidy_type'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['subsidy_type']), axis=1)
         
        df['housing_type'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['housing_type']), axis=1)
          
        df['outcome_date'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['outcome_date']), axis=1)
           
        df['outcome'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['outcome']), axis=1)
        
        df['services_rendered'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['services_rendered']), axis=1)
           
        df['activities'] = df.apply(lambda x: HousingToolBox.RedactForCovid(x['service_type'], x['Post 12/1/21 Elig Date?'], x['activities']), axis=1)
        
        
        def FY22ReleaseRemove(ServiceType, PostTwelveOne, HRARelease):
            if ServiceType == "Advice Only" and PostTwelveOne == "No":
                return "Keep"
            elif ServiceType == "Brief Legal Assistance" and PostTwelveOne == "No":
                return "Keep"
            elif HRARelease == "Yes":
                return "Keep"
            else:
                return "Remove"
        df['Remove post 12-1 Cases?'] = df.apply(lambda x: ReleaseRemove(x['service_type'], x['Post 12/1/21 Elig Date?'], x['HRA Release?']), axis=1)
        
        df = df[df['Remove post 12-1 Cases?'] != 'Remove']
        
        '''
        
        #making excel sheet
        
        borough_dictionary = dict(tuple(df.groupby('Assigned Branch/CC')))

        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                header_format = workbook.add_format({
                'text_wrap':True,
                'bold':True,
                'valign': 'middle',
                'align': 'center'
                })
                
                
                worksheet = writer.sheets[i]
                
                #Add column header data back in
                
                worksheet.freeze_panes(1,0)
                worksheet.set_column('A:BL',20)
                ColShape= int(df.shape[1])
                #worksheet.set_column ('AM:AM',30,link_format)
                HLLoc=df.columns.get_loc("Hyperlinked CaseID#")
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
        output_filename = f.filename

        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)

        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Formatted " + f.filename)
        
       
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>UAHPLP Report Prep</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Prep Cases for UA/HPLP External Report:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=UAC-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1964" target="_blank">HPLP/UAC External Report</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
