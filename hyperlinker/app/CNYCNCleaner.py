#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os
from zipfile import ZipFile

@app.route("/CNYCNCleaner", methods=['GET', 'POST'])
def CNYCNCleaner():
    #upload file from computer
    if request.method == 'POST':
        print(request.files['file'])
        print(request.files.getlist('file')[:])
        print(len(request.files.getlist('file')[:]))
        print(request.form['HOPPdate'])
        print(request.form['CNYCNdate'])
        #print((request.files ['file']))
        #Take date from input field on cleaning tool
        HOPPContractDate = request.form['HOPPdate']
        #Turn HOPP date into an integer
        HOPPContractDate =  DataWizardTools.DateMaker(HOPPContractDate)
        #Take date from input field on cleaning tool
        CNYCNContractDate = request.form['CNYCNdate']
        #Turn CNYCN date into integer
        CNYCNContractDate =  DataWizardTools.DateMaker(CNYCNContractDate)
        #Take end date from input field on cleaning tool
        HOPPEndDate = request.form['HOPPEndDate']
        CNYCNEndDate = request.form['CNYCNEndDate']
        #Turn End Date into an integer
        HEndDateConstruct = DataWizardTools.DateMaker(HOPPEndDate)
        CEndDateConstruct = DataWizardTools.DateMaker(CNYCNEndDate)
        #Return date to a string in CNYCN appropriate format
        HOPPEndDate = datetime.strptime(HOPPEndDate, '%Y-%m-%d').strftime('%m/%d/%Y')
        CNYCNEndDate = datetime.strptime(CNYCNEndDate, '%Y-%m-%d').strftime('%m/%d/%Y') 
        print(HOPPContractDate)
        print(CNYCNContractDate)
        print(HOPPEndDate)
        print(CNYCNEndDate)
        
        
        #adds blank dataframe
        df = pd.DataFrame()
        
        for i in request.files.getlist('file'):
            print (i)
            #turn the excel file into a dataframe, but skip the top 2 rows if they are blank
            test = pd.read_excel(i)
            test.fillna('',inplace=True)
            if test.iloc[0][0] == '':
                tdf = pd.read_excel(i,skiprows=2)
                print("skipped two rows")
            else:
                tdf = pd.read_excel(i)
                print("did not skip")
                
            print(tdf)
           
           #Add dataframes to blank dataframe
            df = df.append(tdf, ignore_index=True)
                      
        '''#assigns varaible f to one file arbitrarily
        f = request.files.getlist('file')[0]
        
        if len(request.files.getlist('file')[:]) == 2:
            g = request.files.getlist('file')[1]
                
        print(f)
        if len(request.files.getlist('file')[:]) == 2:
            print(g)
        #document_list=[request.files['file']]
        #NumOfDocs=len(document_list)
        #print("Number of documents is ", NumOfDocs)
        
                     
        


        if len(request.files.getlist('file')[:]) == 2:
            test2 = pd.read_excel(g)
            test2.fillna('',inplace=True)
            if test2.iloc[0][0] == '':
                df2 = pd.read_excel(g,skiprows=2)
            else:
                df2 = pd.read_excel(g)'''
            
        
        
        #Remove Rows without Case ID values
        df.fillna('',inplace = True)
        df['Matter/Case ID#'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Matter/Case ID#']),axis=1)        
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        

        #Create Hyperlinks
        df['ClientID'] = df['Matter/Case ID#']
        df['Case ID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)     
        
        #Office Abbreviator
        
        df['Assigned Branch/CC'] = df.apply(lambda x: DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis = 1)
        
        #Has Type of Assistance - Done in Conditional Formatting
        #Has Servicer - Done in Conditional Formatting
        #Has reason for distress - Done in Conditional Formatting
        #If closed, has outcome - see if it comes up
        #if mortgage modified, has NewPay? - see if it comes up 
        
        #add benefits column together (even if funds obtained is in crappy format):
        
        def Floatifier(FundsObtained):
            if FundsObtained == '':
                return 0
            elif isinstance(FundsObtained,int) == True or isinstance(FundsObtained,float) == True:
                return FundsObtained
            elif ',' in FundsObtained or '$' in FundsObtained:
                FundsObtained = FundsObtained.replace(',','')
                FundsObtained = FundsObtained.replace('$','')
                try: 
                    FundsObtained = float(FundsObtained)
                    return FundsObtained
                except:
                    print('did not float')
                    return 'Needs Fixing'
            else:
                return 'Needs Fixing'
        
        df['Funds Obtained'] = df.apply(lambda x: Floatifier(x['Funds Obtained']),axis = 1)
        
        def BenefitsSum (FundsObtained, DebtDischarged, SettlementAmount,TotalSavedFromRate,PrincipalReduction):
            
            if isinstance(FundsObtained,int) == True or isinstance(FundsObtained,float) == True:
                BenefitSum = DebtDischarged + SettlementAmount + TotalSavedFromRate + PrincipalReduction + FundsObtained
                return BenefitSum
            elif FundsObtained == 'Needs Fixing':
                return 'Fix Funds Obtained'
            
        df['Benefits'] = df.apply(lambda x: BenefitsSum(x['Funds Obtained'],x['Debt Discharged In Short Sales'],x['Settlement Amount'],x['Total Saved Due To Rate Reduction'],x['Amount Of Principal Reduction']),axis = 1)
        
        #Giving things the CNYCN name:
        
        def FundingSourceNamer(FundsNum,FundsNum2):
            FundsNum = str(FundsNum)
            if FundsNum.startswith('2') == True or FundsNum2.startswith('2') == True:
                return 'HOPP'
            elif FundsNum.startswith('54') == True or FundsNum2.startswith('54') == True:
                return 'CNYCN'
            elif FundsNum.startswith('555') == True or FundsNum2.startswith('555') == True:
                return 'CNYCN'
            elif FundsNum.startswith('556') == True or FundsNum2.startswith('556') == True:
                return 'SENIOR'
            else:   
                return 'Funding Code Error'
                
                
        def EthnicityGuesser(Race,Language):
            if Race == "Latina/o/x" or Race == "Hispanic":
                return 'Hispanic'
            elif Race == 'Other' and Language == 'Spanish':
                return 'Hispanic'
            else:
                return 'Non-Hispanic'
        
        def RaceNamer(Race):
            if Race == "Black/African American/African Descent":
                return "Black/African American"
            elif Race == "White (Not Hispanic)":
                return "White"
            elif Race == "Asian or Pacific Islander":
                return "Asian"
            elif Race == "Hispanic" or Race == "Latina/o/x":
                return "Other"
            else: 
                return Race
        
        def OutcomeReplacer(Outcome):
            if Outcome == "Obtained credit/budget counseling" or Outcome == "Referred to legal services":
                return "Referral"
            elif Outcome == "Resolved non-mortgage lien issue":
                return "Resolved Non-mortgage Lien Issue"
            elif Outcome == "Resolved non-mortgage lien":
                return "Resolved Non-mortgage Lien Issue"
            elif Outcome == "Mortgage Refinanced-In House":
                return "Mortgage Modified - In House"
            else:
                return Outcome
        
        def ModStatusReplacer(ModStatus):
            if ModStatus == "Lender/Servicer Requested Addition Documents":
                return "Lender/Servicer Requested Additional Documents"
            else:
                return ModStatus
        
        def LegalAssistanceReplacer(LegalAssistance):
            if LegalAssistance == "Representation in Good Faith Proceeding":
                return "Litigation"
            else:
                return LegalAssistance
        
        
        def NewPayNoZeroes(NewPay):
            if NewPay == 0:
                return ''
            else:
                return NewPay
        
        def DateTexter (Date):
            Date = str(Date)
            if ":" in Date:
                return '=TEXT("' + Date + '","mm/dd/yyyy")'
            else:   
                return Date
        
        # This function labels cases as in/out of contract period based on date input in cleaning tool   
        #If no date is entered on cleaning tool, no case is deleted and if case is still open, it is not deleted.
        def OldCaseDeleter (FundingSource,DateClosed):
            if DateClosed == "":
                return "In contract"
            elif FundingSource == 'HOPP':
                if HOPPContractDate == "":
                    return "In contract"
                elif DateClosed < HOPPContractDate:
                    return "delete"
                else:
                    return "In contract"
            elif FundingSource != 'HOPP':
                if CNYCNContractDate == "":
                    return "In Contract"
                if DateClosed < CNYCNContractDate:
                    return "delete"
                else:
                    return "In contract"
                    
        #Fill in blank Time Updated with Date Opened
        def TimeUpFiller (TimeUp, DateO):
            if TimeUp == "":
                return DateO
            else:
                return TimeUp
                
        df['Time Updated'] = df.apply(lambda x: TimeUpFiller(x['Time Updated'],x['Date Opened']),axis=1)
                    
        #This function takes any case with a time updated after the end of contract date input to cleaning tool and changes the time updated to match the end of contract date
        def DateBackDater (TimeUpdated, TimeUpdatedConstruct, FundingSource):
            #print(type(TimeUpdated))
            #print(type(EndDate))
            if TimeUpdatedConstruct == "":
                return ""
            elif FundingSource == 'HOPP':
                if HOPPEndDate == "":
                    return TimeUpdated
                elif HEndDateConstruct < TimeUpdatedConstruct:
                    return HOPPEndDate
                else:
                    return TimeUpdated
            elif FundingSource != 'HOPP':
                if CNYCNEndDate == "":
                    return TimeUpdated
                elif CEndDateConstruct < TimeUpdatedConstruct:
                    return CNYCNEndDate
                else:
                    return TimeUpdated
            else:
                return TimeUpdated
                
        
        df['FundingSource'] = df.apply(lambda x: FundingSourceNamer(x['FundsNum'],x['Secondary Funding Codes']),axis = 1)
        df['Time Updated Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Time Updated']),axis = 1) 
        print("Time updated Old")
        print(df['Time Updated'])
        df['Time Updated'] = df.apply(lambda x: DateBackDater(x['Time Updated'],x['Time Updated Construct'],x['FundingSource']), axis = 1)
        print("Time updated New")
        print(df['Time Updated'])
                      
        df['Staff'] = df['Caseworker Name']
        df['IntakeDate'] = df.apply(lambda x: DateTexter(x['Date Opened']),axis = 1)
        df['ServDate'] = df.apply(lambda x: DateTexter(x['Time Updated']),axis = 1)
        df['Race'] = df.apply(lambda x: RaceNamer(x['Race (CNYCN)']),axis = 1)
        df['Ethnicity'] = df.apply(lambda x: EthnicityGuesser(x['Race (CNYCN)'],x['Language']),axis = 1)
        df['Language'] = ''
        df['Children'] = df['Number of People under 18']
        df['Adults'] = df['Number of People 18 and Over']
        df['Seniors'] = df['Number Of Seniors In Household']
        df['Income'] = df['Total Annual Income ']
        df['Household'] = ''
        df['ZIP'] = df['Zip Code']
        df['County'] = df['County of Residence']
        df['PrimaryDist'] = df['FPU Prim Src Client Prob']
        df['SecondaryDist'] = df['FPU Sec Src Client Prob']
        df['Units'] = df['FPU Num Prop Units']
        df['PurchaseDate'] = ''
        df['OrigDate'] = ''
        df['OrigDate2'] = ''
        df['OrigAmount'] = ''
        df['OrigAmount2'] = ''
        df['OrigTerm'] = ''
        df['OrigTerm2'] = ''
        df['LoanOwner'] = ''
        df['LoanOwner2'] = ''
        df['IntakePrincipal'] = ''
        df['IntakePrincipal2'] = ''
        df['IntakeProd'] = ''
        df['IntakeProd2'] = ''
        df['IntakeRate'] = ''
        df['IntakeRate2'] = ''
        df['IntakePay'] = ''
        df['IntakePay2'] = ''
        df['InterestOnly'] = ''
        df['InterestOnly2'] = ''
        df['LoanStatus'] = ''
        df['LoanStatus2'] = ''
        df['LPDate'] = ''
        df['Servicer'] = df['Servicer']
        df['LoanNumber'] = ''
        df['Servicer2'] = df['Servicer2']
        df['LoanNumber2'] = ''
        df['Violation'] = ''
        df['Violation2'] = ''
        df['ServicerChange'] = ''
        df['ServicerChange2'] = ''
        df['Conference#'] = ''
        df['FirstConference'] = ''
        df['BadFaith'] = ''
        df['LegalHr'] = ''
        df['PrimaryLegal'] = df.apply(lambda x: LegalAssistanceReplacer(x['Type Of Assistance']),axis=1)
        df['SecondLegal'] = df.apply(lambda x: LegalAssistanceReplacer(x['Secondary Assistance Type']),axis=1)
        df['PrimarySandy'] = ''
        df['SecondSandy'] = ''
        df['ModStatus'] = df.apply(lambda x: ModStatusReplacer(x['Loan Modification Status']),axis =1)
        df['ModStatus2'] = df.apply(lambda x: ModStatusReplacer(x['Loan Modification Status 2']),axis =1)
        df['PrimaryOutcome'] = df.apply(lambda x: OutcomeReplacer(x['FPU Primary Outcome']),axis = 1)
        df['SecondOutcome'] = df.apply(lambda x: OutcomeReplacer(x['FPU Secondary Outcome']),axis = 1)
        df['PrimaryOutcome2'] = ''
        df['SecondaryOutcome2'] = ''
        df['NewPrincipal'] = ''
        df['NewPrincipal2'] = ''
        df['NewTerm'] = ''
        df['NewTerm2'] = ''
        df['NewType'] = ''
        df['NewType2'] = ''
        df['NewRate'] = ''
        df['NewRate2'] = ''
        df['NewPay'] = df.apply(lambda x: NewPayNoZeroes(x['FPU Mod PITI Payment - 1st']),axis = 1)
        df['NewPay2'] = df.apply(lambda x: NewPayNoZeroes(x['FPU Mod PITI Payment - 2nd']),axis = 1)
        df['PrincipalForgive'] = ''
        df['PrincipalForgive2'] = ''
        df['ForbearAmt'] = ''
        df['ForbearAmt2'] = ''
        df['SSPrice'] = ''
        df['SSDate'] = ''
        df['Forgiven'] = ''
        df['DILDate'] = ''
        #df['Benefits'] = 'Slightly More Complicated - see above'
        df['NewHousing'] = ''
        df['CaseClose'] = ''
        #Turn Date Closed dates into integers
        df['Date Closed Construct'] = df.apply(lambda x: DataWizardTools.DateMaker(x['Date Closed']),axis = 1)
        #Use the (integer) dates and the funding source to determine if each case is in the contract
        df['In Contract?'] = df.apply(lambda x: OldCaseDeleter(x['FundingSource'],x['Date Closed Construct']),axis = 1)
        df['Active Cases'] = df.apply(lambda x: OldCaseDeleter(x['FundingSource'],x['Time Updated Construct']),axis = 1)
        print(df['In Contract?'])
        print(df['Active Cases'])

        #Putting everything in the right order
        df = df.sort_values(by=['Caseworker Name'])
        df = df.sort_values(by=['FundsNum'])
        
        #split good for formatting
        df['Branch&Report'] = df['Assigned Branch/CC'] + df['FundingSource']
        
        #Wacky Apostrophes coming out of LegalServer
        df = df.replace("’","'",regex = True)
        
        #Remove cases w date closed before contract date
        df = df[df['In Contract?'] != "delete"]
        
        #Remove Unreportables (servicer, Servdate, primary distress, primarylegal)
        
        def Unreportables(Servicer,Servdate,PrimaryDistress,PrimaryLegal):
            if Servicer == '' or Servdate == '' or PrimaryDistress == '' or PrimaryLegal == '':
                return 'Unreportable'
        
        if request.form.get('remover'):
            df['Unreportable'] = df.apply(lambda x: Unreportables(x['Servicer'],x['ServDate'],x['PrimaryDist'],x['PrimaryLegal']),axis=1)
            df = df[df['Unreportable'] != "Unreportable"]
            
            #print(request.form['CaseChecker'])
            CaseCheckerList = list(request.form['CaseChecker'].split(', '))
            print(CaseCheckerList)
            for i in CaseCheckerList:
                if df ['ClientID'].str.contains(i).any():
                    print(i + " in dataframe, remover on")
                else:
                    print(i + " not in dataframe, remover on")
        
        #Formatting Version
        
        if request.form.get('formatter'):
            
            df = df[df['Active Cases'] != "delete"]
            
            df = df[['FundingSource','ClientID','Staff','IntakeDate','ServDate','Race','Ethnicity','Language','Children','Adults','Seniors','Income','Household','ZIP','County','PrimaryDist','SecondaryDist','Units','PurchaseDate','OrigDate','OrigDate2','OrigAmount','OrigAmount2','OrigTerm','OrigTerm2','LoanOwner','LoanOwner2','IntakePrincipal','IntakePrincipal2','IntakeProd','IntakeProd2','IntakeRate','IntakeRate2','IntakePay','IntakePay2','InterestOnly','InterestOnly2','LoanStatus','LoanStatus2','LPDate','Servicer','LoanNumber','Servicer2','LoanNumber2','Violation','Violation2','ServicerChange','ServicerChange2','Conference#','FirstConference','BadFaith','LegalHr','PrimaryLegal','SecondLegal','PrimarySandy','SecondSandy','ModStatus','ModStatus2','PrimaryOutcome','SecondOutcome','PrimaryOutcome2','SecondaryOutcome2','NewPrincipal','NewPrincipal2','NewTerm','NewTerm2','NewType','NewType2','NewRate','NewRate2','NewPay','NewPay2','PrincipalForgive','PrincipalForgive2','ForbearAmt','ForbearAmt2','SSPrice','SSDate','Forgiven','DILDate','Benefits','NewHousing','CaseClose','Branch&Report']]
            
            #print(request.form['CaseChecker'])
            CaseCheckerList = list(request.form['CaseChecker'].split(', '))
            print(CaseCheckerList)
            for i in CaseCheckerList:
                if df['ClientID'].str.contains(i).any():
                    print(i + " in dataframe, formatter on")
                else:
                    print(i + " not in dataframe, formatter on")
            
        #Cleanup Version
        else:
        
            df['Unreportable'] = df.apply(lambda x: Unreportables(x['Servicer'],x['ServDate'],x['PrimaryDist'],x['PrimaryLegal']),axis=1)
        
            df = df[['Case ID#','Caseworker Name','Time Updated','Client First Name','Client Last Name','Type Of Assistance','FPU Prim Src Client Prob','Servicer','FPU Primary Outcome','FPU Secondary Outcome','Loan Modification Status','FPU Mod PITI Payment - 1st','Benefits','Funds Obtained','FundsNum','Assigned Branch/CC','Date Opened','Race (CNYCN)','Number of People 18 and Over','Number of People under 18','Number Of Seniors In Household','Total Annual Income ','Zip Code','County of Residence','FPU Sec Src Client Prob','FPU Num Prop Units','Secondary Assistance Type','Loan Modification Status 2','FPU Secondary Outcome','FPU Mod PITI Payment - 2nd','Debt Discharged In Short Sales','Settlement Amount','Total Saved Due To Rate Reduction','Amount Of Principal Reduction']]
            
            #print(request.form['CaseChecker'])
            CaseCheckerList = list(request.form['CaseChecker'].split(', '))
            print(CaseCheckerList)
            for i in CaseCheckerList:
                if df['Case ID#'].str.contains(i).any():
                    print(i + " in dataframe, cleanup on")
                else:
                    print(i + " not in dataframe, cleanup on")
                
            '''if df['Case ID#'].str.contains(CaseChecker).any():
                print(str(CaseChecker) + " in dataframe")
            else:   
                print(str(CaseChecker) + " not in dataframe")'''
        
            
        #define function, with single variable of dictionary created immediately above
        def save_xls(ExcelSplit, TabSplit):
            #Create dictionary of dataframes, each named after each unique value of 'assigned branch/cc'
            excel_dict_df = dict(tuple(df.groupby(ExcelSplit)))
            #use 'ZipFile' create empty zip folder, and assign 'newzip' as function-calling name
            with ZipFile("app\\sheets\\zipped\\SplitBoroughs.zip","w") as newzip:
                
                #starts cycling through each dataframe (each borough's data)
                for i in excel_dict_df:
                    
                    #because this is in for loop it creates a new excel file for each 'i' (i.e. each borough)
                    writer = pd.ExcelWriter(path = "app\\sheets\\zipped\\" + i + ".xlsx", engine = 'xlsxwriter')
                    
                    #create a dictionary of dataframes for each unique advocate, within the borough that our 'i' for loop is cycling through
                    tab_dict_df = dict(tuple(excel_dict_df[i].groupby(TabSplit)))
                    
                    #start cycling through each of these new advocate-based dataframes
                    for j in tab_dict_df:
                        
                        #write a tab in the borough's excel sheet, composed just of the advocate's cases, with the advocate's name
                        if request.form.get('formatter'):
                            tab_dict_df[j].to_csv(path_or_buf = "app\\sheets\\zipped\\" + j + ".csv", index = False, sep=str(','))
                        else:
                            tab_dict_df[j].to_excel(writer, j, index = False)
                        
                            #creates ability to format tabs
                            worksheet = writer.sheets[j]
                            
                            #creates and adds formatting to tab
                            workbook = writer.book
                            link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                            problem_format = workbook.add_format({'bg_color':'yellow'})
                    
                            #worksheet = writer.sheets[i]
                            #Type of Assistance, FPU Prim Src Client Prob, Servicer
                            FGHRowRange='F2:H'+str(tab_dict_df[j].shape[0]+1)
                            print(FGHRowRange)
                            
                            #Funds Obtained
                            NRowRange='K1:K'+str(tab_dict_df[j].shape[0]+1)
                            
                            worksheet.set_column('A:A',12,link_format)
                            worksheet.set_column('B:B',20)
                            worksheet.set_column('C:E',15)
                            worksheet.set_column('F:L',27)
                            worksheet.set_column('M:O',14)
                            worksheet.set_column('P:AH',0)
                            worksheet.conditional_format(NRowRange,{'type': 'text',
                                                         'criteria': 'containing',
                                                         'value': 'Fix',
                                                         'format': problem_format})
                            worksheet.conditional_format(FGHRowRange,{'type': 'blanks',
                                                                     'format': problem_format})
                            worksheet.freeze_panes(1,1)
                        
                    #save's excel file
                    writer.save()
                    #adds excel file to zipped folder
                    if request.form.get('formatter'):
                        newzip.write("app\\sheets\\zipped\\" + i + ".csv",arcname = i + ".csv")
                    else:
                        newzip.write("app\\sheets\\zipped\\" + i + ".xlsx",arcname = i + ".xlsx")
          
        #Preparing Excel Document
        if request.form.get('formatter'):
            save_xls(ExcelSplit = 'Branch&Report', TabSplit = 'Branch&Report')
        else:
            save_xls(ExcelSplit = 'Assigned Branch/CC', TabSplit = 'Caseworker Name')

        '''def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                
                worksheet = writer.sheets[i]
                EFGRowRange='E2:G'+str(dict_df[i].shape[0]+1)
                print(EFGRowRange)
                
                KRowRange='K1:K'+str(dict_df[i].shape[0]+1)
                
                if request.form.get('formatter'):
                    worksheet.set_column('A:CE',20)
                else:
                    worksheet.set_column('A:A',12,link_format)
                    worksheet.set_column('B:B',20)
                    worksheet.set_column('C:E',15)
                    worksheet.set_column('F:L',27)
                    worksheet.set_column('M:O',14)
                    worksheet.set_column('P:AH',0)
                    worksheet.conditional_format(KRowRange,{'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'Fix',
                                                 'format': problem_format})
                    worksheet.conditional_format(EFGRowRange,{'type': 'blanks',
                                                             'format': problem_format})
                    worksheet.freeze_panes(1,1)
            writer.save()
        output_filename = "Cleaned Foreclosure Report.xlsx"

        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)'''
        
        output_filename = "Cleaned Foreclosure Report.xlsx"
        if request.form.get('formatter'):
            FilePrefix = "Formatted "
            #return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = FilePrefix + output_filename)
            return send_from_directory('sheets\\zipped','SplitBoroughs.zip', as_attachment = True)
        else:
            FilePrefix = "Cleanup "
            return send_from_directory('sheets\\zipped','SplitBoroughs.zip', as_attachment = True)
        
        #return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = FilePrefix + output_filename)
        #return send_from_directory('sheets\\zipped','SplitBoroughs.zip', as_attachment = True,attachment_filename = FilePrefix + output_filename)
        
        #***#
       
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>CNYCN Cleaner</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Clean Cases for CNYCN</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file multiple =""><input type=submit value=Clean!>
    
    </br>
    </br>
       
    <input type = "date" id="HOPPdate" name="HOPPdate">
    <label for = "HOPPdate"> choose HOPP contract start date! </label>
    
    </br>
    
    <input type = "date" id="HOPPEndDate" name="HOPPEndDate">
    <label for = "HOPPEndDate"> choose HOPP reporting period end date! </label>
    
    </br>
    
    <input type = "date" id="CNYCNdate" name="CNYCNdate">
    <label for = "CNYCNdate"> choose CNYCN contract start date! </label>
    
    </br>
    
    <input type = "date" id="CNYCNEndDate" name="CNYCNEndDate">
    <label for = "CNYCNEndDate"> choose CNYCN reporting period end date! </label>
    
    </br>
    </br>

    <input type="checkbox" id="formatter" name="formatter" value="formatter">
    <label for="formatter"> Format for Submission</label><br>
    
    <input type="checkbox" id="remover" name="remover" value="remover">
    <label for="remover"> Remove Unreportables for Formatted Report</label><br>
    
    </br>
     <input type = "text" id="CaseChecker" name="CaseChecker" input style="width:120%" 
    <label for ="CaseChecker"> <span></span> Type case list</label><br><br>
    
    
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool will help prepare cases for CNYCN Reports.</li>
    <li>This tool is meant to be used in conjunction with the LegalServer report (DOR) called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2328" target="_blank">HOPP/CNYCN Monthly Interim-Central Team</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
