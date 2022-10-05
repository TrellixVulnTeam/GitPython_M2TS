from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools, HousingToolBox
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
#from playsound import playsound



@app.route("/UAHPLPConsolidatedBoroughCleaner", methods=['GET', 'POST'])
def UAHPLPConsolidatedBoroughCleaner():
    if request.method == 'POST':
    
        
    
        print(request.files['file'])
        f = request.files['file']
        
        test = pd.read_excel(f)
        
        test.fillna('',inplace=True)
        
        #Preparing Excel Document
        
        
        
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
        
        df['Assigned Branch/CC'] = df.apply(lambda x : DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis=1)   

        if request.form.get('MLS'):
            print('MLS Version')
            df = df[df['Assigned Branch/CC'] == "MLS"]
        if request.form.get('BkLS'):
            print('BkLS Version')
            df = df[df['Assigned Branch/CC'] == "BkLS"]
        
        
        #Has to have a Housing Type of Case
        

        #Has to have a Housing Level of Service 

        #Level of Service is HOLD FOR REVIEW tester? incorporate into above

        #Eligiblity date tester - blank or not? *changed for 12/1/21 dates
       
        df['EligDateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['HAL Eligibility Date']), axis=1)
        
        
        
        df['OpenedDateConstruct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Opened']), axis=1)
       

        #PA Tester if theres no dhci, not needed for post-covid advice/brief cases *pre 12/1 cases don't need income verification document
        def PATester (PANum,DHCI):

            if DHCI == "DHCI Form" and PANum == "":
                return "Not Needed due to DHCI"
            else:
                return PANum
            
        df['Gen Pub Assist Case Number'] = df.apply(lambda x: PATester(x['Gen Pub Assist Case Number'],x['Housing Income Verification']),axis = 1)
        
        #Outcome Tester - date no outcome or outcome no date
        
        def OutcomeTester(Outcome,OutcomeDate):
            if OutcomeDate != "" and Outcome == "":
                return "Needs Outcome"
            else:
                return Outcome
                
        df['Housing Outcome'] = df.apply(lambda x: OutcomeTester(x['Housing Outcome'],x['Housing Outcome Date']),axis = 1)
                
        def OutcomeDateTester(Outcome,OutcomeDate):
            if Outcome != "" and OutcomeDate == "":
                return "Needs Outcome Date"
            else:
                return OutcomeDate      

        df['Housing Outcome Date'] = df.apply(lambda x: OutcomeDateTester(x['Housing Outcome'],x['Housing Outcome Date']),axis = 1)
        
        
        #Test if a case is actually a housing case
        def NonHousingCase (LegalProblemCode):
            if LegalProblemCode.startswith('6') == True:
                return ''
            else:
                return 'Non-Housing Case, Please Review'
                
        df['Housing Case?'] = df.apply(lambda x: NonHousingCase(x['Legal Problem Code']),axis = 1)
        
        #ERAP Testing
        df['ERAP Tester'] = df.apply(lambda x: HousingToolBox.ERAPTester(x['Housing Type Of Case'],x['ERAP Involved Case?'],x['Stayed ERAP Case?'],x['Is stayed ERAP case active?']), axis = 1)
            
    
        #Is everything okay with a case? Also remove if Eligdate is from prior year

        def TesterTester (EligConstruct,HRARelease,HousingLevel,HousingType,EligDate,PANum,Outcome,OutcomeDate,ERAP,HousingCase):
           
            if EligConstruct != '' and EligConstruct < 20220701 :
                return 'Eligibility date from prior contract year'
            elif HRARelease == "" or HRARelease == "No" or HRARelease == " ":
                return 'Case Needs Attention'
            elif HousingLevel == "" or HousingLevel == "Hold For Review":
                return 'Case Needs Attention'
            elif HousingType == '':
                return 'Case Needs Attention'
            elif EligDate == '':
                return 'Case Needs Attention'
            elif PANum == '':
                return 'Case Needs Attention'
            elif Outcome == 'Needs Outcome' or OutcomeDate == 'Needs Outcome Date':
                return 'Case Needs Attention'
            elif 'Needs' in ERAP or 'needs' in ERAP:
                return 'Case Needs Attention'
            elif HousingCase == 'Non-Housing Case, Please Review':
                return 'Case Needs Attention'
            else:
                return 'No Cleanup Necessary'
            
        df['Tester Tester'] = df.apply(lambda x: TesterTester(x['EligDateConstruct'],x['HRA Release?'],x['Housing Level of Service'],x['Housing Type Of Case'],x['HAL Eligibility Date'],x['Gen Pub Assist Case Number'],x['Housing Outcome'],x['Housing Outcome Date'],x['ERAP Tester'],x['Housing Case?']),axis=1)
        
        
        #Delete if everything's okay **

        df = df[df['Tester Tester'] == "Case Needs Attention"]

        #sort by case handler
        
        df = df.sort_values(by=['Primary Advocate'])
        
        #playsound("app\\static\\sound.wav")
        #print ("played")
        
        #Create borough-specific tabs for MLS and BkLS as needed
        if request.form.get('MLS'):
            
            Evelyn_Casehandlers = ['Delgadillo, Omar','Heller, Steven E','Latterner, Matt J','Tilyayeva, Rakhil','Almanzar, Yocari']
            Diana_V_Casehandlers = ['Abbas, Sayeda','Hao, Lindsay','He, Ricky','Spencer, Eleanor G','Wilkes, Nicole','Allen, Sharette','Ortiz, Matthew B','Sun, Dao','Risener, Jennifer A','Surface, Ben L','Velasquez, Diana']
            Diana_G_Casehandlers = ['Saxton, Jonathan G','Orsini, Mary K','Duffy-Greaves, Kevin','Freeman, Daniel A','Gokhale, Aparna S','Gonzalez, Matias','Gonzalez, Matias G','Labossiere, Samantha J.','Shah, Ami Mahendra']
            #Keiannis_Casehandlers = ['Almanzar, Milagros','Briggs, John M','Dittakavi, Archana','Gonzalez-Munoz, Rossana G','Honan, Thomas J','James, Lelia','Kelly, Kitanya','Yamasaki, Emily Woo J','McCune, Mary','Vogltanz, Amy K','Whedan, Rebecca','McDonald, John']
            Dennis_Casehandlers = ['Braudy, Erica','Kulig, Jessica M','Mercedes, Jannelys J','Harshberger, Sae','Black, Rosalind','Gelly-Rahim, Jibril']
            Rosa_Casehandlers = ['Acron, Denise D','Anunkor, Ifeoma O','Reyes, Nicole V','Vega, Rita']
            Anthony_Casehandlers = ['Basu, Shantonu J','Arboleda, Heather M','Grater, Ashley P','Sharma, Sagar','Evers, Erin C.','Frierson, Jerome C','Rockett, Molly C']

            Access_Line = ["Pierre, Haenley","Ortega, Luis","Djourab, Atteib","Suriel, Sal","Villanueva, Anthony","Ruiz-Caceres, Gaby A","Yeh, Victoria","Paz, Alex","Khanam, Aysha"]
            
            def IntakeAssign(Casehandler, Advocate):
                if Casehandler == 'Vergeli, Evelyn':
                    return "Evelyn V."
                elif Casehandler == 'Velasquez, Diana':
                    return "Diana V."
                elif Casehandler == 'Garcia, Diana':
                    return "Diana G."
                elif Casehandler == 'Trinidad, Ayla A':
                    return "Ayla T."
                #elif Casehandler == 'De Jesus, Christine' or Casehandler == 'Garcia, Keiannis':
                #    return "Christine D."
                elif Casehandler == 'Sanchez, Dennis':
                    return "Dennis S."
                elif Casehandler == 'Acosta, Rosa F':
                    return "Rosa A."
                elif Casehandler == 'Benitez, Vicenta':
                    return "Vincenta B."
                elif Casehandler == 'Garcia, Delci T.':
                    return "Delci G"
                elif Casehandler == 'Garcia, Alexandra A.':
                    return "Alexandra G."
                elif Casehandler == 'Villanueva, Anthony':
                    return "Tony V."
                elif Advocate in Evelyn_Casehandlers:
                    return "Evelyn V."
                elif Advocate in Diana_V_Casehandlers:
                    return "Diana V."
                elif Advocate in Diana_G_Casehandlers:
                    return "Diana G."
                #elif Advocate in Keiannis_Casehandlers:
                #    return "Christine D."
                elif Advocate in Dennis_Casehandlers:
                    return "Dennis S."
                elif Advocate in Rosa_Casehandlers:
                    return "Rosa A."
                elif Advocate in Anthony_Casehandlers:
                    return "Tony V."
                elif Casehandler in Access_Line:
                    global whoseturn
                    try: 
                        whoseturn += 1
                    except:
                        whoseturn = 1
                    if whoseturn == 7:
                        whoseturn = 1
                        
                    if whoseturn == 1:
                        return 'Evelyn V.'
                    elif whoseturn == 2:
                        return 'Diana V.'
                    elif whoseturn == 3:
                        return 'Diana G.'
                    elif whoseturn == 4:
                        return 'Ayla T.'
                    elif whoseturn == 5:
                        return 'Dennis S.'
                    elif whoseturn == 6:
                        return 'Tony V.'
                else:
                    return "zzMiscellaneous"

            df['Assigned Paralegal'] = df.apply(lambda x: IntakeAssign(x['Caseworker Name'], x['Primary Advocate']),axis = 1)
        
        
        elif request.form.get('BkLS'):
            def IntakeAssign(Casehandler):
                if Casehandler == 'Wong, Angela':
                    return 'Angela Wong'
                elif Casehandler == 'Lane, Diane':
                    return 'Diane Lane'
                elif Casehandler == 'Oquendo, Joann':
                    return 'Joann Oquendo'
                elif Casehandler == 'Mullen, Evan M':
                    return 'Evan Mullen'
                elif Casehandler == 'Spivey, Joseph':
                    return 'Joseph Spivey'
                elif Casehandler == 'Moss, Julieta':
                    return 'Julieta Moss'
                else:
                    global whoseturn
                    try: 
                        whoseturn += 1
                    except:
                        whoseturn = 1
                    if whoseturn == 7:
                        whoseturn = 1
                        
                    if whoseturn == 1:
                        return 'Angela Wong'
                    elif whoseturn == 2:
                        return 'Diane Lane'
                    elif whoseturn == 3:
                        return 'Joann Oquendo'
                    elif whoseturn == 4:
                        return 'Evan Mullen'
                    elif whoseturn == 5:
                        return 'Joseph Spivey'
                    elif whoseturn == 6:
                        return 'Julieta Moss'
                
                

            df['Assigned Paralegal'] = df.apply(lambda x: IntakeAssign(x['Caseworker Name']),axis = 1)
        
        else:
            df['Assigned Paralegal'] = 'No Assignment Made'

        df['Intake User'] = df['Caseworker Name']
        #Put everything in the right order
        
        df = df[['Hyperlinked CaseID#',
        'Primary Advocate',
        'Date Opened',
        'Date Closed',
        "Client First Name",
        "Client Last Name",
        
        "HRA Release?",
        "Housing Level of Service",
        "Housing Type Of Case",
        "HAL Eligibility Date",
        
        
        
        "Gen Pub Assist Case Number",
        
        "Housing Income Verification",
        #"Housing Signed DHCI Form",
        
        "Housing Outcome",
        "Housing Outcome Date",
        "Housing Case?",
        "ERAP Involved Case?",
        "Stayed ERAP Case?",
        "Is stayed ERAP case active?",
        "ERAP Tester",
        "Gen Case Index Number", 
        "Tester Tester",
        "Assigned Branch/CC",
        "Assigned Paralegal",
        "Intake User",
        "Percentage of Poverty"
        ]]      
        
        
        '''
        Graveyard
                  _(_)_                          wWWWw   _
      @@@@       (_)@(_)   vVVVv     _     @@@@  (___) _(_)_
     @@()@@ wWWWw  (_)\    (___)   _(_)_  @@()@@   Y  (_)@(_)
      @@@@  (___)     `|/    Y    (_)@(_)  @@@@   \|/   (_)\
       /      Y       \|    \|/    /(_)    \|      |/      |
    \ |     \ |/       | / \ | /  \|/       |/    \|      \|/
jgs \\|//   \\|///  \\\|//\\\|/// \|///  \\\|//  \\|//  \\\|// 
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


        
        df['Post 12/1/21 Elig Date?'] = df.apply(lambda x: HousingToolBox.PostTwelveOne(x['EligDateConstruct']), axis=1)
        
        def FY22ReleaseTester(HRARelease,LevelOfService,IndexNum,PostTwelveOne):
            LevelOfService = str(LevelOfService)
            IndexNum = str(IndexNum)
            if PostTwelveOne == "No":
                if LevelOfService.startswith("Advice") or LevelOfService.startswith("Brief"):
                    if IndexNum == '' or IndexNum.startswith('N') == True or IndexNum.startswith('n') == True:
                        return 'Unnecessary due to Limited Service'
                    else:
                        return HRARelease
                else:
                    return HRARelease
            else:
                return HRARelease
                
        df['HRA Release?'] = df.apply(lambda x: ReleaseTester(x['HRA Release?'],x["Housing Level of Service"],x['Gen Case Index Number'],x['Post 12/1/21 Elig Date?']),axis = 1)
        
        #PA Tester if theres no dhci, not needed for post-covid advice/brief cases *pre 12/1 cases don't need income verification document
        def FY22PATester (PANum,DHCI,PostTwelveOne,LevelOfService,EligDate,OpenDate):
            LevelOfService = str(LevelOfService)
            if PostTwelveOne == "No" and LevelOfService.startswith("Advice") and EligDate != '':
                return "Unnecessary due limited service"
            elif PostTwelveOne == "No" and LevelOfService.startswith("Brief") and EligDate != '':
                return "Unnecessary due limited service"         
            elif DHCI == "DHCI Form" and PANum == "":
                return "Not Needed due to DHCI"
            else:
                return PANum
        
        '''
        
        
        
        #Split into different tabs
        
        if request.form.get('MLS'):
            allgood_dictionary = dict(tuple(df.groupby('Assigned Paralegal')))
        elif request.form.get('BkLS'):
            allgood_dictionary = dict(tuple(df.groupby('Assigned Paralegal')))
        else:
            allgood_dictionary = dict(tuple(df.groupby('Assigned Branch/CC')))
        
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                ws = writer.sheets[i]
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                regular_format = workbook.add_format({'font_color':'black'})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
                ws.autofilter('B1:ZZ1')
                ws.freeze_panes(1, 2)
                GKRowRange='G1:K'+str(dict_df[i].shape[0]+1)
                print(GKRowRange)
                #get dynamic conditional formatting that looks for column header, not column location
                #(first_row, first_col, last_row, last_col)
                shape = (dict_df[i].shape[0]+1)
                ERAPLoc=df.columns.get_loc("ERAP Tester")


                ws.conditional_format(GKRowRange,{'type': 'blanks',
                                                 'format': problem_format})                                                 
                ws.conditional_format('H2:H100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Hold For Review',
                                                 'format': problem_format})
                ws.conditional_format('G2:G100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'No',
                                                 'format': problem_format})
                ws.conditional_format('M2:N100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
                ws.conditional_format('O2:O100000',{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Review',
                                                 'format': problem_format})
                ws.conditional_format(0,ERAPLoc,shape,ERAPLoc,{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Needs',
                                                 'format': problem_format})
                ws.conditional_format(0,ERAPLoc,shape,ERAPLoc,{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'needs',
                                                 'format': problem_format})
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = allgood_dictionary, path = "app\\sheets\\" + output_filename)
       
        if request.form.get('MLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "MLS " + f.filename)
        elif request.form.get('BkLS'):
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "BkLS " + f.filename)
        else:
            return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)
        
        

    return '''
    <!doctype html>
    <title>UAHPLP Consolidated Borough Cleaner</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>UAHPLP Consolidated Borough Cleaner:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    
    </br>
    </br>
    <input type="checkbox" id="MLS" name="MLS" value="MLS">
    <label for="MLS"> MLS Compliance</label><br>
    <input type="checkbox" id="BkLS" name="BkLS" value="BkLS">
    <label for="BkLS"> BkLS Compliance</label><br>
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2369" target="_blank">HPLP/UAC Internal Report All Cases</a>.</li>
    
   
    </br>
    <a href="/">Home</a>
    '''
