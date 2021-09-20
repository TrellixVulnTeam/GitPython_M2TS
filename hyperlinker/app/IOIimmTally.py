from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, ImmigrationToolBox, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
from datetime import date
import pandas as pd
#import textwrap


@app.route("/IOIimmTally", methods=['GET', 'POST'])
def upload_IOIimmTally():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        test = pd.read_excel(f)
        
        test.fillna('',inplace=True)
        
        #date definition?
        today = date.today()
        print(today.month)
        
        if today.month >= 8:
            howmanymonths = today.month - 7
        else:
            howmanymonths = today.month + 12 - 7
        
        print(howmanymonths)
        
        #Cleaning
        if test.iloc[0][0] == '':
            df = pd.read_excel(f,skiprows=2)
        else:
            df = pd.read_excel(f)
        
        df.fillna('',inplace=True)
       #Create Hyperlinks
        df['Hyperlinked Case #'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)
        
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Bronx Legal Services','BxLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Brooklyn Legal Services','BkLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Queens Legal Services','QLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Manhattan Legal Services','MLS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Staten Island Legal Services','SILS')
        df['Assigned Branch/CC'] = df['Assigned Branch/CC'].str.replace('Legal Support Unit','LSU')
        
        
        
        #Determining 'level of service' from 3 fields       
          
        df['HRA Level of Service'] = df.apply(lambda x: ImmigrationToolBox.HRA_Level_Service(x['Close Reason'],x['Level of Service']), axis=1)
        
        
       #HRA Case Coding
       #Putting Cases into HRA's Baskets!  
        df['HRA Case Coding'] = df.apply(lambda x: ImmigrationToolBox.HRA_Case_Coding(x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA Level of Service'],x['IOI Does Client Have A Criminal History? (IOI 2)']), axis=1)

        #Dummy SLPC for Juvenile Cases
        def DummySLPC(LPC,SLPC):
            LPC = str(LPC)
            SLPC = str(SLPC)
            if LPC == "44 Minor Guardianship / Conservatorship" or LPC == "42 Neglected/Abused/Dependent":
                return 'N/A'
            else:
                return SLPC
                
        df['Special Legal Problem Code'] = df.apply(lambda x: DummySLPC(x['Legal Problem Code'],x['Special Legal Problem Code']), axis=1)
    
        df['HRA Service Type'] = df['HRA Case Coding'].apply(lambda x: x[:2] if x != 'Hold For Review' else '')

        df['HRA Proceeding Type'] = df['HRA Case Coding'].apply(lambda x: x[3:] if x != 'Hold For Review' else '')
        
        #Giving things better names in cleanup sheet
        
        df['DHCI form?'] = df['Has Declaration of Household Composition and Income (DHCI) Form?']
        
        df['Consent form?'] = df['IOI HRA Consent Form? (IOI 2)']
        
        df['Client Name'] = df['Full Person/Group Name (Last First)']
        
        df['Office'] = df['Assigned Branch/CC']
        
        df['Country of Origin'] = df['IOI Country Of Origin (IOI 1 and 2)']
        
        df['Substantial Activity'] = df['IOI Substantial Activity (Choose One)']
        
        df['Date of Substantial Activity'] = df['Custom IOI Date substantial Activity Performed']
        
        #Income Waiver
        
        def Income_Exclude(IncomePct,Waiver,Referral):   
            IncomePct = int(IncomePct)
            Waiver = str(Waiver)
            if Referral == 'Action NY':
                return ''
            elif IncomePct > 200 and Waiver.startswith('2') == False:
                return 'Needs Income Waiver'
            else:
                return ''

        df['Exclude due to Income?'] = df.apply(lambda x: Income_Exclude(x['Percentage of Poverty'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)'],x['IOI Referral Source (IOI 2)']), axis=1)
        
        #Eligibility_Date & Rollovers 
        
        def Eligibility_Date(Substantial_Activity_Date,Effective_Date,Date_Opened):
            if Substantial_Activity_Date != '':
                return Substantial_Activity_Date
            elif Effective_Date != '':
                return Effective_Date
            else:
                return Date_Opened
        df['Eligibility_Date'] = df.apply(lambda x : Eligibility_Date(x['IOI Date Substantial Activity Performed 2022'],x['IOI HRA Effective Date (optional) (IOI 2)'],x['Date Opened']), axis = 1)
        
        #Manipulable Dates               
        
        
        df['Open Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Eligibility_Date']),axis = 1)
        
        df['Subs Month'] = df['IOI Date Substantial Activity Performed 2022'].apply(lambda x: str(x)[5:7])
        df['Subs Day'] = df['IOI Date Substantial Activity Performed 2022'].apply(lambda x: str(x)[8:])
        df['Subs Year'] = df['IOI Date Substantial Activity Performed 2022'].apply(lambda x: str(x)[:4])
        df['Subs Construct'] = df['Subs Year'] + df['Subs Month'] + df['Subs Day']
        df['Subs Construct'] = df.apply(lambda x : x['Subs Construct'] if x['Subs Construct'] != '' else 0, axis = 1)
        
        df['Outcome1 Month'] = df['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[:2])
        df['Outcome1 Day'] = df['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[3:5])
        df['Outcome1 Year'] = df['IOI Outcome 2 Date (IOI 2)'].apply(lambda x: str(x)[6:])
        df['Outcome1 Construct'] = df['Outcome1 Year'] + df['Outcome1 Month'] + df['Outcome1 Day']

        df['Outcome2 Month'] = df['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[:2])
        df['Outcome2 Day'] = df['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[3:5])
        df['Outcome2 Year'] = df['IOI Secondary Outcome Date 2 (IOI 2)'].apply(lambda x: str(x)[6:])
        df['Outcome2 Construct'] = df['Outcome2 Year'] + df['Outcome2 Month'] + df['Outcome2 Day']       
        
        
        #DHCI Form
                
        def DHCI_Needed(DHCI,Open_Construct,LoS):
            if Open_Construct == '':
                return ''
            elif LoS.startswith('Advice'):
                return ''
            elif LoS.startswith('Brief'):
                return ''
            elif int(Open_Construct) < 20180701:
                return ''
            elif DHCI != 'Yes':
                return 'Needs DHCI Form'
            else:
                return ''
        
        df['Needs DHCI?'] = df.apply(lambda x: DHCI_Needed(x['Has Declaration of Household Composition and Income (DHCI) Form?'],x['Open Construct'],x['Level of Service']), axis=1)
        
        
        
        #Needs Substantial Activity to Rollover into FY'20
        
              
         #Needs Substantial Activity to Rollover into FY'21
         
         #Needs Substantial Activity to Rollover into FY'22
        
        
        def Needs_Rollover(Open_Construct,Substantial_Activity, Substantial_Activity_Date,CaseID) :
            if int(Open_Construct) >= 20210701:
                return ''
            elif Substantial_Activity != '' and int(Substantial_Activity_Date) >20210701 and int(Substantial_Activity_Date) <=20220630:
                return ''
            elif CaseID in ImmigrationToolBox.ReportedFY19 or CaseID in ImmigrationToolBox.ReportedFY21 or CaseID in ImmigrationToolBox.ReportedFY20:
                return 'Needs Substantial Activity in FY22'
            else: return ''
        df['Needs Substantial Activity?'] = df.apply(lambda x: Needs_Rollover(x['Open Construct'],x['IOI FY22 Substantial Activity 2022'],x['Subs Construct'],x['Matter/Case ID#']), axis=1)  
        
        


        #Outcomes
        
                
        #if there are two outcomes choose which outcome to report based on which one happened more recently 
                
        def OutcomeToReport (Outcome1,OutcomeDate1,Outcome2,OutcomeDate2,ServiceLevel,CloseDate):
            if OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Advice':
                return 'Advice given'
            elif OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Brief Service':
                return 'Advice given'
            elif CloseDate != '' and ServiceLevel == 'Full Rep or Extensive Service' and Outcome1 == '' and Outcome2 == '':
                return '*Needs Outcome*'
            elif OutcomeDate1 >= OutcomeDate2:
                return Outcome1
            elif OutcomeDate2 > OutcomeDate1:
                return Outcome2
            else:
                return 'no actual outcome'
        
        df['Outcome To Report'] = df.apply(lambda x: OutcomeToReport(x['IOI Outcome 2 (IOI 2)'],x['Outcome1 Construct'],x['IOI Secondary Outcome 2 (IOI 2)'],x['Outcome2 Construct'],x['HRA Level of Service'],x['Date Closed']), axis=1)
      
        #make it add the outcome date as well (or tell you if you need it!)        

        def OutcomeDateToReport (Outcome1,OutcomeDate1,Outcome2,OutcomeDate2,ServiceLevel,CloseDate,ActualOutcomeDate1,ActualOutcomeDate2):
            if OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Advice':
                return CloseDate
            elif OutcomeDate1 == '' and OutcomeDate2 == '' and CloseDate != '' and ServiceLevel == 'Brief Service':
                return CloseDate           
            elif OutcomeDate1 >= OutcomeDate2:
                return ActualOutcomeDate1
            elif OutcomeDate2 > OutcomeDate1:
                return ActualOutcomeDate2
            else:
                return '*Needs Outcome Date*'
            
                
        df['Outcome Date To Report'] = df.apply(lambda x: OutcomeDateToReport(x['IOI Outcome 2 (IOI 2)'],x['Outcome1 Construct'],x['IOI Secondary Outcome 2 (IOI 2)'],x['Outcome2 Construct'],x['HRA Level of Service'],x['Date Closed'],x['IOI Outcome 2 Date (IOI 2)'],x['IOI Secondary Outcome Date 2 (IOI 2)']), axis=1)
        

        #kind of glitchy - if it has an outcome date but no outcome it doesn't say *Needs Outcome*
        
        #add LSNYC to start of case numbers 
        
        df['Unique_ID'] = 'LSNYC'+df['Matter/Case ID#']
        
        #take second letters of first and last names
        
        df['Last_Initial'] = df['Client Last Name'].str[1]
        df['First_Initial'] = df['Client First Name'].str[1]

        #Year of birth
        df['Year_of_Birth'] = df['Date of Birth'].str[-4:]
        
        #Unique Client ID#
        df['Unique Client ID#'] = df['First_Initial'] + df['Last_Initial'] + df['Year_of_Birth']       
        
        #Deliverable Categories
        
        def Deliverable_Category(HRA_Coded_Case,Income_Cleanup,Age_at_Intake,ClientsNames):
            if Income_Cleanup == 'Needs Income Waiver':
                return 'Needs Cleanup'
            elif HRA_Coded_Case == 'T2-RD' and Age_at_Intake <= 21:
                return 'Tier 2 (minor removal)'
            elif HRA_Coded_Case == 'T2-RD' and ClientsNames in ImmigrationToolBox.AtlasClientsNames :
                return 'Tier 2 (minor removal)'
            elif HRA_Coded_Case == 'T2-RD':
                return 'Tier 2 (removal)'
            elif HRA_Coded_Case.startswith('T2') == True:
                return 'Tier 2 (other)'
            elif HRA_Coded_Case.startswith('T1')== True:
                return 'Tier 1'
            elif HRA_Coded_Case.startswith('B') == True:
                return 'Brief'
            else:
                return 'Needs Cleanup'

        df['Deliverable Tally'] = df.apply(lambda x: Deliverable_Category(x['HRA Case Coding'],x['Exclude due to Income?'],x['Age at Intake'], x['Client Name']), axis=1)
        
        #make all cases for any client that has a minor removal tally, into also being minor removal cases
        
        
        df['Modified Deliverable Tally'] = ''
        
        dfs = df.groupby('Unique Client ID#',sort = False)

        tdf = pd.DataFrame()
        for x, y in dfs:
            for z in y['Deliverable Tally']:
                if z == 'Tier 2 (minor removal)':
                    y['Modified Deliverable Tally'] = 'Tier 2 (minor removal)'
            tdf = tdf.append(y)
        df = tdf
        
        
        #write function to identify blank 'modified deliverable tallies' and add it back in as the original deliverable tally
        df.fillna('',inplace= True)

        def fillBlanks(ModifiedTally,Tally):
            if ModifiedTally == '':
                return Tally
            else:
                return ModifiedTally
        df['Modified Deliverable Tally'] = df.apply(lambda x: fillBlanks(x['Modified Deliverable Tally'],x['Deliverable Tally']),axis=1)

        #Reportable?
        
        df['Reportable?'] = df.apply(lambda x: ImmigrationToolBox.ReportableTester(x['Exclude due to Income?'],x['Needs DHCI?'],x['Needs Substantial Activity?'],x['Deliverable Tally']),axis=1)
        
        #***add code to make it so that it deletes any extra 'brief' cases for clients that have mutliple cases
        
        
        
        
        
        #gender
        def HRAGender (gender):
            if gender == 'Male' or gender == 'Female':
                return gender
            else:
                return 'Other'
        df['Gender'] = df.apply(lambda x: HRAGender(x['Gender']), axis=1)
        
        #county=borough
        df['Borough'] = df['County of Residence']
        
        #household size etc.
        df['Household_Size'] = df['Number of People under 18'].astype(int) + df['Number of People 18 and Over'].astype(int)
        df['Number_of_Children'] = df['Number of People under 18']
        
        #Income Eligible?
        df['Annual_Income'] = df['Total Annual Income ']
        def HRAIncElig (PercentOfPoverty):
            if PercentOfPoverty > 200:
                return 'NO'
            else:
                return 'YES'
        df['Income_Eligible'] = df.apply(lambda x: HRAIncElig(x['Percentage of Poverty']), axis=1)
        
        def IncWaiver (eligible,waiverdate):
            if eligible == 'NO' and waiverdate != '':
                return 'Income'
            else:
                return ''
        df['Waiver_Type'] = df.apply(lambda x: IncWaiver(x['Income_Eligible'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)']), axis=1)
        
        def IncWaiverDate (waivertype,date):
            if waivertype != '':
                return date
            else:  
                return ''
                
        df['Waiver_Approval_Date'] = df.apply(lambda x: IncWaiverDate(x['Waiver_Type'],x['IOI HRA WAIVER APPROVAL DATE if over 200% of FPL (IOI 2)']), axis = 1)
        
        #Referrals
        def Referral (referral):
            if referral == "Action NY":
                return "ActionNYC"
            elif referral == "HRA":
                return "HRA-DSS"
            elif referral == "Other":
                return "Other"
            elif referral == "":
                return "None"
            else:
                return ""
                
        df['Referral_Source'] = df.apply(lambda x: Referral(x['IOI Referral Source (IOI 2)']), axis = 1)
        
        #Pro Bono Involvement
        def ProBonoCase (branch, pai):
            if branch == "LSU" or pai == "Yes":
                return "YES"
            else:
                return "NO"
                
        df['Pro_Bono'] = df.apply(lambda x:ProBonoCase(x['Assigned Branch/CC'], x['PAI Case?']), axis = 1)
        
        #Prior Enrollment
        
        def PriorEnrollment (casenumber):
            if casenumber in ImmigrationToolBox.ReportedFY21:
                return 'FY 21'
            elif casenumber in ImmigrationToolBox.ReportedFY20:
                return 'FY 20'
            elif casenumber in ImmigrationToolBox.ReportedFY19:
                return 'FY 19'
                
                
        df['Prior_Enrollment_FY'] = df.apply(lambda x:PriorEnrollment(x['Matter/Case ID#']), axis = 1)
        
        #Other Cleanup
        df['Service_Type_Code'] = df['HRA Service Type']
        df['Proceeding_Type_Code'] = df['HRA Proceeding Type']
        df['Outcome'] = df['Outcome To Report']
        df['Outcome_Date'] = df['Outcome Date To Report']
        df['Seized_at_Border'] = df['IOI Was client apprehended at border? (IOI 2&3)']
        df['Group'] = ''
          
        #Construct Summary Tables
        #city_pivot = pd.pivot_table(df,index=['Office'],values=['Deliverable Tally'],aggfunc='count',fill_value=0)
        city_pivot = pd.pivot_table(df,values=['Unique_ID'], index=['Office'], columns=['Modified Deliverable Tally'], aggfunc=lambda x: len(x.unique()))
        
        
        city_pivot.reset_index(inplace=True)
        print(city_pivot)
        #city_pivot['Unique_ID','Annual Goal']='Test' 

        #adds the totals for each column
        #city_pivot['Unique_ID',[city_pivot.loc['Total','Brief':'Tier 2 (removal)'] = city_pivot.sum(axis=0)]] 
        #city_pivot.loc[city_pivot['Unique_ID',['Total','Brief'] = city_pivot.sum(axis=0)]]
        #inter['total'] = inter.sum(axis=0, numeric_only= True)
        #cells.hideColumn(1,4)
        #city_pivot.style.hide_columns([1,4])
        
        #Add totals to all columns
        #city_pivot.loc['I2':'X8'] = city_pivot.sum(axis=0) 
        #city_pivot.sum(axis=1, skipna=None, level=0, numeric_only=None, min_count=0)
        #df.groupby(level=1).sum().
        #city_pivot.groupby('Brief').sum()
        #print(city_pivot.loc['I2':'X8'] = city_pivot.sum(axis=0) )
        #city_pivot.loc['Total'] = city_pivot.groupby(level=0).sum(-3)
        #print(city_pivot)
        #totalBrief=city_pivot.groupby(level=[0][0]['J']).sum()
        #print(totalBrief)
        city_pivot.loc['Total'] = city_pivot.sum(axis=0)
        print(city_pivot)

        #Add Goals to Summary Tables:
                        
        def BriefGoal(Office):
            if Office == "BxLS":
                return 20
            elif Office == "BkLS":
                return 20
            elif Office == "MLS":
                return 20
            elif Office == "QLS":
                return 20
            elif Office == "SILS":
                return 20
            elif Office == "LSU":
                return 140
            else:
                return 240
        
        def Tier1Goal(Office):
            if Office == "BxLS":
                return 34
            elif Office == "BkLS":
                return 43
            elif Office == "MLS":
                return 38
            elif Office == "QLS":
                return 34
            elif Office == "SILS":
                return 15
            elif Office == "LSU":
                return 100
            else:
                return 264
                
        def T2MRGoal(Office):
            if Office == "BxLS":
                return 62
            elif Office == "BkLS":
                return 62
            elif Office == "MLS":
                return 63
            elif Office == "QLS":
                return 62
            elif Office == "SILS":
                return 63
            elif Office == "LSU":
                return 62
            else:
                return 374
        def T2OTHRGoal(Office):
            if Office == "BxLS":
                return 7
            elif Office == "BkLS":
                return 11
            elif Office == "MLS":
                return 2
            elif Office == "QLS":
                return 7
            elif Office == "SILS":
                return 25
            elif Office == "LSU":
                return 92
            else:
                return 144
        def T2RmvlGoal(Office):
            if Office == "BxLS":
                return 24
            elif Office == "BkLS":
                return 24
            elif Office == "MLS":
                return 24
            elif Office == "QLS":
                return 24
            elif Office == "SILS":
                return 24
            elif Office == "LSU":
                return 0
            else:
                return 120
                                      
        city_pivot.reset_index(inplace=True)

        if request.form.get('Proportional'):   
            print('Hello proportional!')
        
        #Add goals to City Pivot and percentage calculators, alternating w Proportional Goals   
            city_pivot['Unique_ID','Brief '] = city_pivot['Unique_ID','Brief']
            city_pivot['Unique_ID','Ann Brf Goal'] = city_pivot.apply(lambda x: BriefGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann Brf %']=city_pivot['Unique_ID','Brief']/city_pivot['Unique_ID','Ann Brf Goal']
            city_pivot['Unique_ID','Prop Brf Goal'] = round((city_pivot['Unique_ID','Ann Brf Goal']/12 * howmanymonths),2)
            city_pivot['Unique_ID','Prop Brf %']=city_pivot['Unique_ID','Brief ']/city_pivot['Unique_ID','Prop Brf Goal']
            city_pivot['Unique_ID','Tier 1 '] = city_pivot['Unique_ID','Tier 1']
            city_pivot['Unique_ID','Ann T1 Goal'] = city_pivot.apply(lambda x: Tier1Goal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T1 %']=city_pivot['Unique_ID','Tier 1']/city_pivot['Unique_ID','Ann T1 Goal']
            city_pivot['Unique_ID','Prop T1 Goal'] = round((city_pivot['Unique_ID','Ann T1 Goal']/12 * howmanymonths),2)
            city_pivot['Unique_ID','Prop T1 %']=city_pivot['Unique_ID','Tier 1 ']/city_pivot['Unique_ID','Prop T1 Goal']
            city_pivot['Unique_ID','T2 MR'] = city_pivot['Unique_ID','Tier 2 (minor removal)']
            city_pivot['Unique_ID','Ann T2 MR Goal'] = city_pivot.apply(lambda x: T2MRGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T2 MR %']=city_pivot['Unique_ID','T2 MR']/city_pivot['Unique_ID','Ann T2 MR Goal']
            city_pivot['Unique_ID','Prop T2MR Goal'] = round((city_pivot['Unique_ID','Ann T2 MR Goal']/12 * howmanymonths),2)
            city_pivot['Unique_ID','Prop T2MR %']=city_pivot['Unique_ID','T2 MR']/city_pivot['Unique_ID','Prop T2MR Goal']
            city_pivot['Unique_ID','T2 Other'] = city_pivot['Unique_ID','Tier 2 (other)']
            city_pivot['Unique_ID','Ann T2O Goal'] = city_pivot.apply(lambda x: T2OTHRGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T2O %']=city_pivot['Unique_ID','T2 Other']/city_pivot['Unique_ID','Ann T2O Goal']
            city_pivot['Unique_ID','Prop T2O Goal'] = round((city_pivot['Unique_ID','Ann T2O Goal']/12 * howmanymonths),2)
            city_pivot['Unique_ID','Prop T2O %']=city_pivot['Unique_ID','T2 Other']/city_pivot['Unique_ID','Prop T2O Goal']
            city_pivot['Unique_ID','T2 Rmvl'] = city_pivot['Unique_ID','Tier 2 (removal)']
            city_pivot['Unique_ID','Ann T2 Rmvl Goal'] = city_pivot.apply(lambda x: T2RmvlGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T2 Rmvl %']=city_pivot['Unique_ID','T2 Rmvl']/city_pivot['Unique_ID','Ann T2 Rmvl Goal']
            city_pivot['Unique_ID','Prop T2R Goal'] = round((city_pivot['Unique_ID','Ann T2 Rmvl Goal']/12 * howmanymonths),2)
            city_pivot['Unique_ID','Prop T2R %']=city_pivot['Unique_ID','T2 Rmvl']/city_pivot['Unique_ID','Prop T2R Goal']  
        
        else:
            print('Annual here')
            #Add goals to City Pivot and percentage calculators, alternating - No Proportional Goals
            city_pivot['Unique_ID','Brief '] = city_pivot['Unique_ID','Brief']
            city_pivot['Unique_ID','Ann Brf Goal'] = city_pivot.apply(lambda x: BriefGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann Brf %']=city_pivot['Unique_ID','Brief']/city_pivot['Unique_ID','Ann Brf Goal']
            #city_pivot['Unique_ID','Prop Brf Goal'] = round((city_pivot['Unique_ID','Ann Brf Goal']/12 * howmanymonths),2)
            #city_pivot['Unique_ID','Prop Brf %']=city_pivot['Unique_ID','Brief ']/city_pivot['Unique_ID','Prop Brf Goal']
            city_pivot['Unique_ID','Tier 1 '] = city_pivot['Unique_ID','Tier 1']
            city_pivot['Unique_ID','Ann T1 Goal'] = city_pivot.apply(lambda x: Tier1Goal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T1 %']=city_pivot['Unique_ID','Tier 1']/city_pivot['Unique_ID','Ann T1 Goal']
            #city_pivot['Unique_ID','Prop T1 Goal'] = round((city_pivot['Unique_ID','Ann T1 Goal']/12 * howmanymonths),2)
            #city_pivot['Unique_ID','Prop T1 %']=city_pivot['Unique_ID','Tier 1 ']/city_pivot['Unique_ID','Prop T1 Goal']
            city_pivot['Unique_ID','T2 MR'] = city_pivot['Unique_ID','Tier 2 (minor removal)']
            city_pivot['Unique_ID','Ann T2 MR Goal'] = city_pivot.apply(lambda x: T2MRGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T2 MR %']=city_pivot['Unique_ID','T2 MR']/city_pivot['Unique_ID','Ann T2 MR Goal']
            #city_pivot['Unique_ID','Prop T2MR Goal'] = round((city_pivot['Unique_ID','Ann T2 MR Goal']/12 * howmanymonths),2)
            #city_pivot['Unique_ID','Prop T2MR %']=city_pivot['Unique_ID','T2 MR']/city_pivot['Unique_ID','Prop T2MR Goal']
            city_pivot['Unique_ID','T2 Other'] = city_pivot['Unique_ID','Tier 2 (other)']
            city_pivot['Unique_ID','Ann T2O Goal'] = city_pivot.apply(lambda x: T2OTHRGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T2O %']=city_pivot['Unique_ID','T2 Other']/city_pivot['Unique_ID','Ann T2O Goal']
            #city_pivot['Unique_ID','Prop T2O Goal'] = round((city_pivot['Unique_ID','Ann T2O Goal']/12 * howmanymonths),2)
            #city_pivot['Unique_ID','Prop T2O %']=city_pivot['Unique_ID','T2 Other']/city_pivot['Unique_ID','Prop T2O Goal']
            city_pivot['Unique_ID','T2 Rmvl'] = city_pivot['Unique_ID','Tier 2 (removal)']
            city_pivot['Unique_ID','Ann T2 Rmvl Goal'] = city_pivot.apply(lambda x: T2RmvlGoal(x['Office','']), axis=1)
            city_pivot['Unique_ID','Ann T2 Rmvl %']=city_pivot['Unique_ID','T2 Rmvl']/city_pivot['Unique_ID','Ann T2 Rmvl Goal']
            #city_pivot['Unique_ID','Prop T2R Goal'] = round((city_pivot['Unique_ID','Ann T2 Rmvl Goal']/12 * howmanymonths),2)
            #city_pivot['Unique_ID','Prop T2R %']=city_pivot['Unique_ID','T2 Rmvl']/city_pivot['Unique_ID','Prop T2R Goal']  
		

        #Add zeros to blank cells:
        city_pivot.fillna(0, inplace=True)
            
              
        #REPORTING VERSION Put everything in the right order
        df = df[['Hyperlinked Case #','Office','Primary Advocate','Client Name','Special Legal Problem Code','Level of Service','Needs DHCI?','Exclude due to Income?','Needs Substantial Activity?','Country of Origin','Outcome To Report','Modified Deliverable Tally','Reportable?']]
                   
        #Removed - 'Unique_ID','Last_Initial','First_Initial','Year_of_Birth','Gender','Country of Origin','Borough','Zip Code','Language','Household_Size','Number_of_Children','Annual_Income','Income_Eligible','Waiver_Type','Waiver_Approval_Date','Eligibility_Date','Referral_Source','Service_Type_Code','Proceeding_Type_Code','Outcome','Outcome_Date','Seized_at_Border','Group','Prior_Enrollment_FY','Pro_Bono','Special Legal Problem Code','HRA Level of Service','HRA Case Coding','IOI FY22 Substantial Activity 2022','IOI Date Substantial Activity Performed 2022','IOI Was client apprehended at border? (IOI 2&3)','Deliverable Tally',
        
                
        
        #Preparing Excel Document
        
        #Bounce to Excel
        borough_dictionary = dict(tuple(df.groupby('Office')))
        
        #dimensions = worksheet.shape
        #print(dimensions)
        #len(('worksheet','MLS'))
        #worksheet.info()
        #total_rows = MLS.count
        #print (total_rows)
        #len((worksheet.index)
        #worksheet.shape[0]  
        #df.shape[0]
        #count_row = df.shape[0]  # Gives number of rows   
        #print(count_row)
        #print(borough_dictionary['MLS'].shape[0])
           
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                print(dict_df[i].shape[0]+1)
                city_pivot.to_excel(writer, sheet_name='City Pivot',index=True,header = False,startrow=1)
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                #cols = [textwrap.wrap(x, width=20) for x in cols] - First attempt at text wrapping
                percent_format = workbook.add_format({})#({'bold':True})
                percent_format.set_num_format('0%')
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                officehead_format = workbook.add_format({'font_color':'black','bold':True})
                header_format = workbook.add_format({'text_wrap':True,'bold':True,'valign': 'top','align': 'center','bg_color' : '#eeece1'})
                HeaderBlock_format = workbook.add_format({'text_wrap':True,'bold':True,'valign': 'top','align': 'center','bg_color' : '#d2b48c'})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                totals_format = workbook.add_format({'bold':True})
                border_format=workbook.add_format({'border':1,'align':'left','font_size':10})
                CityPivot = writer.sheets['City Pivot']
                worksheet = writer.sheets[i]
                #print(dict_df[i].shape[1])
                #HeaderRange='A1:'str(dict_df[i].shape[1]'+'1')
                #print(HeaderRange)
                worksheet.autofilter('A1:M1')
                worksheet.set_column('A:A',20,link_format)
                worksheet.set_column('B:B',19)
                worksheet.set_column('C:BL',30)
                worksheet.freeze_panes(1,1)
                worksheet.write_comment('E1', 'All Cases Must Have a Special Legal Problem Code')
                
                #Add column header data back in
                for col_num, value in enumerate(city_pivot['Unique_ID'].columns.values):
                    CityPivot.write(1, col_num +3, value, header_format)
                    
                if request.form.get('Proportional'):   
                    print('Hello proportional!')
                                    
                    #CityPivot sheet formatting w Proportional Goals
                    CityPivot.write('C2', 'Office',officehead_format)
                    CityPivot.write('C9', 'Citywide',officehead_format)
                    CityPivot.write('J2:AH2', 'Brief ',header_format)
                    #JOTYD
                    CityPivot.write('J2', 'Brief ',HeaderBlock_format)
                    CityPivot.write('O2', 'Tier 1 ',HeaderBlock_format)
                    CityPivot.write('T2', 'T2 MR',HeaderBlock_format)
                    CityPivot.write('Y2', 'T2 Other',HeaderBlock_format)
                    CityPivot.write('AD2', 'T2 Rmvl',HeaderBlock_format)
                    CityPivot.freeze_panes(0,3)
                    CityPivot.set_column('A:B',0)
                    CityPivot.set_column('C:C',6,officehead_format)
                    CityPivot.set_column('D:I',0)
                    CityPivot.set_column('J:V',11)
                    CityPivot.set_column('W:W',11)
                    CityPivot.set_row(0,0)
                    #CityPivot.set_row(2,0) no longer need to hide this row, it was deleted when the headers disappeared
                    CityPivot.set_row(1,31)
                    #CityPivot.set_row(1,31,header_format) removed to prevent beige bormatting to infinitum
                    #l,n,q,s,v,x,a,c,f,h
                    CityPivot.set_column('L:L',11,percent_format)
                    CityPivot.set_column('N:N',11,percent_format)
                    CityPivot.set_column('Q:Q',11,percent_format)
                    CityPivot.set_column('S:S',11,percent_format)
                    CityPivot.set_column('V:V',11,percent_format)
                    CityPivot.set_column('X:X',11,percent_format)
                    CityPivot.set_column('AA:AA',11,percent_format)
                    CityPivot.set_column('AC:AC',11,percent_format)
                    CityPivot.set_column('AF:AF',11,percent_format)
                    CityPivot.set_column('AH:AH',11,percent_format)
                    CityPivot.conditional_format( 'C2:AH9' , { 'type' : 'cell' ,'criteria': '!=','value':'""','format' : border_format} )
                    #CityPivot.set_row(6,20,totals_format)
                    #CityPivot.set_column('D:D',20,percent_format)
                    #CityPivot.set_column('F:F',20,percent_format)
                    
                    #CityPivot.write('A7', 'Totals', totals_format)
                    #CityPivot.set_column('L:L',10,percent_format)
                
                else:
                    print('Annual here')
                
                    #CityPivot sheet formatting - No Proportional Goals
                    CityPivot.write('C2', 'Office',officehead_format)
                    CityPivot.write('C9', 'Citywide',officehead_format)
                    CityPivot.write('J2:AH2', 'Brief ',header_format)
                    #JMPSV
                    CityPivot.write('J2', 'Brief ',HeaderBlock_format)
                    CityPivot.write('M2', 'Tier 1 ',HeaderBlock_format)
                    CityPivot.write('P2', 'T2 MR',HeaderBlock_format)
                    CityPivot.write('S2', 'T2 Other',HeaderBlock_format)
                    CityPivot.write('V2', 'T2 Rmvl',HeaderBlock_format)
                    CityPivot.freeze_panes(0,3)
                    CityPivot.set_column('A:B',0)
                    CityPivot.set_column('C:C',6,officehead_format)
                    CityPivot.set_column('D:I',0)
                    CityPivot.set_column('J:V',11)
                    CityPivot.set_column('W:W',14)
                    CityPivot.set_row(0,0)
                    #CityPivot.set_row(2,0) no longer need to hide this row, it was deleted when the headers disappeared
                    CityPivot.set_row(1,31)
                    #CityPivot.set_row(1,31,header_format)
                    CityPivot.set_column('L:L',11,percent_format)
                    CityPivot.set_column('O:O',11,percent_format)
                    CityPivot.set_column('R:R',11,percent_format)
                    CityPivot.set_column('U:U',11,percent_format)
                    CityPivot.set_column('X:X',13,percent_format)
                    CityPivot.conditional_format( 'C2:AH9' , { 'type' : 'cell' ,'criteria': '!=','value':'""','format' : border_format} )
                    #CityPivot.set_row(6,20,totals_format)
                    #CityPivot.set_column('D:D',20,percent_format)
                    #CityPivot.set_column('F:F',20,percent_format)
                    
                    #CityPivot.write('A7', 'Totals', totals_format)
                    #CityPivot.set_column('L:L',10,percent_format)
                
                
                EFRowRange='E1:F'+str(dict_df[i].shape[0]+1)
                print(EFRowRange)
                JRowRange='J1:J'+str(dict_df[i].shape[0]+1)
                print(JRowRange)
             
                worksheet.conditional_format(EFRowRange,{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
                '''worksheet.conditional_format('F1:F100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})'''
                worksheet.conditional_format('F1:F100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Hold for Review"',
                                                 'format': problem_format})                                                 
                worksheet.conditional_format('G1:G100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs DHCI Form"',
                                                 'format': problem_format})
                worksheet.conditional_format('H1:H100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Income Waiver"',
                                                 'format': problem_format})
                worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Substantial Activity in FY22"',
                                                 'format': problem_format})
                worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Substantial Activity Date"',
                                                 'format': problem_format})
                worksheet.conditional_format(JRowRange,{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
                worksheet.conditional_format('K1:K100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"*Needs Outcome*"',
                                                 'format': problem_format})
                worksheet.conditional_format('L1:L100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Cleanup"',
                                                 'format': problem_format})
                worksheet.conditional_format('M1:M100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"No"',
                                                 'format': problem_format})
                
            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)
           

        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleanup " + f.filename)
        
    return '''
    <!doctype html>
    <title>IOI Immigration Tally</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Tally your IOI Immigration Cases and Case Cleanup:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=IOI-ify!>
    
    </br>
    </br>
    <input type="checkbox" id="Proportional" name="Proportional" value="Proportional">
    <label for="Proportional"> Proportional Deliverable Targets</label><br>
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1918" target="_blank">"Grants Management IOI 2 (3459) Report"</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for IOI.</li> 
    <li>Once you have identified this file, click ‘IOI-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
