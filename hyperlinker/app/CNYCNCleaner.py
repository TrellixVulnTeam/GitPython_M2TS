#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db, DataWizardTools
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os

@app.route("/CNYCNCleaner", methods=['GET', 'POST'])
def CNYCNCleaner():
    #upload file from computer
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
        df['ClientID'] = df['Matter/Case ID#']
        df['Case ID#'] = df.apply(lambda x : DataWizardTools.Hyperlinker(x['Matter/Case ID#']),axis=1)     
        
        #Office Abbreviator
        
        df['Assigned Branch/CC'] = df.apply(lambda x: DataWizardTools.OfficeAbbreviator(x['Assigned Branch/CC']),axis = 1)
        
        #Has Type of Assistance - Done in Conditional Formatting
        #Has Servicer - Done in Conditional Formatting
        #Has reason for distress - Done in Conditional Formatting
        #If closed, has outcome - see if it comes up
        #if mortgage modified, has NewPay? - see if it comes up 
        
        
        
        #Giving things the CNYCN name:
        
        def FundingSourceNamer(FundsNum):
            FundsNum = str(FundsNum)
            if FundsNum.startswith('2') == True:
                return 'HOPP'
            elif FundsNum.startswith('5') == True:
                return 'CNYCN'
            else:   
                return 'Funding Code Error'
                
        def EthnicityGuesser(Race,Language):
            if Race == "Latina/o/x" or Race == "Hispanic":
                return 'Hispanic'
            elif Race == 'Other' and Language == 'Spanish':
                return 'Hispanic'
            else:
                return 'Non-Hispanic'
                
        df['FundingSource'] = df.apply(lambda x: FundingSourceNamer(x['FundsNum']),axis = 1)
        df['Staff'] = df['Caseworker Name']
        df['IntakeDate'] = df['Date Opened']
        df['ServDate'] = df['Time Updated']
        df['Race'] = df['Race (CNYCN)']
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
        df['Servicer2'] = ''
        df['LoanNumber2'] = ''
        df['Violation'] = ''
        df['Violation2'] = ''
        df['ServicerChange'] = ''
        df['ServicerChange2'] = ''
        df['Conference#'] = ''
        df['FirstConference'] = ''
        df['BadFaith'] = ''
        df['LegalHr'] = ''
        df['PrimaryLegal'] = df['Type Of Assistance']
        df['SecondLegal'] = df['Secondary Assistance Type']
        df['PrimarySandy'] = ''
        df['SecondSandy'] = ''
        df['ModStatus'] = df['Loan Modification Status']
        df['ModStatus2'] = df['Loan Modification Status 2']
        df['PrimaryOutcome'] = df['FPU Primary Outcome']
        df['SecondOutcome'] = df['FPU Secondary Outcome']
        
        
        """
        ***DO the rest of the Sheet
        df[''] = ''
        df[''] = ''
        df[''] = ''
        df[''] = ''
        df[''] = df['']
        df[''] = df['']
        df[''] = df['']
        df[''] = df['']
        """
        
        
        #Putting everything in the right order
        df = df.sort_values(by=['Caseworker Name'])
        df = df.sort_values(by=['FundsNum'])
        
        
        #Formatting Version
        
        if request.form.get('formatter'):
            
            #***separate HOPP from CNYCN (by sort or sheet)
            df = df[['FundingSource','ClientID','Staff','IntakeDate','ServDate','Race','Ethnicity','Language','Children','Adults','Seniors','Income','Household','ZIP','County','PrimaryDist','SecondaryDist','Units','PurchaseDate','OrigDate','OrigDate2','OrigAmount','OrigAmount2','OrigTerm','OrigTerm2','LoanOwner','LoanOwner2','IntakePrincipal','IntakePrincipal2','IntakeProd','IntakeProd2','IntakeRate','IntakeRate2','IntakePay','IntakePay2','InterestOnly','InterestOnly2','LoanStatus','LoanStatus2','LPDate','Servicer','LoanNumber','Servicer2','LoanNumber2','Violation','Violation2','ServicerChange','ServicerChange2','Conference#','FirstConference','BadFaith','LegalHr','PrimaryLegal','SecondLegal','PrimarySandy','SecondSandy','ModStatus','ModStatus2','PrimaryOutcome','SecondOutcome','Assigned Branch/CC']]
            
        #Cleanup Version
        else:
        
            df = df[['Case ID#','Caseworker Name','Client First Name','Client Last Name','Type Of Assistance','FPU Prim Src Client Prob','Servicer','FPU Primary Outcome','Loan Modification Status','FPU Mod PITI Payment - 1st','Assigned Branch/CC','FundsNum','Date Opened','Time Updated','Race (CNYCN)','Number of People 18 and Over','Number of People under 18','Number Of Seniors In Household','Total Annual Income ','Zip Code','County of Residence','FPU Sec Src Client Prob','FPU Num Prop Units','Secondary Assistance Type','Loan Modification Status 2','FPU Secondary Outcome','FPU Mod PITI Payment - 2nd','Funds Obtained','Debt Discharged In Short Sales','Settlement Amount','Total Saved Due To Rate Reduction','Amount Of Principal Reduction']]
        
        
        
        
        #Preparing Excel Document

        borough_dictionary = dict(tuple(df.groupby('Assigned Branch/CC')))

        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                problem_format = workbook.add_format({'bg_color':'yellow'})
                
                worksheet = writer.sheets[i]
                
                if request.form.get('formatter'):
                    worksheet.set_column('A:CC',20)
                else:
                    worksheet.set_column('A:A',12,link_format)
                    worksheet.set_column('B:B',20)
                    worksheet.set_column('C:D',15)
                    worksheet.set_column('E:J',30)
                    worksheet.set_column('K:AF',0)

                    worksheet.conditional_format('E2:G100000',{'type': 'blanks',
                                                             'format': problem_format})
                    worksheet.freeze_panes(1,1)
            writer.save()
        output_filename = f.filename

        save_xls(dict_df = borough_dictionary, path = "app\\sheets\\" + output_filename)

        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleanup " + f.filename)
        
        #***#
       
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>CNYCN Cleaner</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Clean Cases for CNYCN</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    
    </br>
    </br>

    <input type="checkbox" id="formatter" name="formatter" value="formatter">
    <label for="formatter"> Format for Submission</label><br>
    
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This should will help prepare cases for CNYCN Reports.</li>
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2328" target="_blank">HOPP/CNYCN Monthly Interim-Central Team</a>.</li>
    
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
