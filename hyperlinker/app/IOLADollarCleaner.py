#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/IOLADollarCleaner", methods=['GET', 'POST'])
def IOLADollarCleaner():
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
        if 'Matter/Case ID#' not in df.columns:
            df['Matter/Case ID#'] = df['id']
        
        df.fillna('',inplace=True)
        #MAKE THIS WORK!!
        def NoIDDelete(CaseID):
            if CaseID == '' or CaseID == 'nan':
                return 'No Case ID'
            else:
                return str(CaseID)
        df['Matter/Case ID#'] = df.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        df = df[df['Matter/Case ID#'] != 'No Case ID']
        
        
        
        last7 = df['Matter/Case ID#'].apply(lambda x: x[-7:])
        df['Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + df['Matter/Case ID#'] +'"' +')'
        move = df['Hyperlinked Case #']
        del df['Hyperlinked Case #']
        df.insert(0,'Hyperlinked Case #',move)           
        del df['Matter/Case ID#']
        
        #Test if monthly amount is in Closing Screen
        
        def DAPMonthlyTester (DAPSSIMonthly,DAPSSDMonthly,ClosingMonthly):
            ClosingMonthly = float(ClosingMonthly[1:].replace(",",""))
            DAPSSIMonthly = float(DAPSSIMonthly[1:].replace(",",""))
            DAPSSDMonthly = float(DAPSSDMonthly[1:].replace(",",""))
            
            if DAPSSIMonthly + DAPSSDMonthly > ClosingMonthly + 100:
                return "Monthly Award in Closing Screen Should be as Large as in DAP Screen"
            else:
                return ''
        df['DAP Monthly $ Tester'] = df.apply(lambda x : DAPMonthlyTester(x['DAP Monthly XVI -- SSI'],x['DAP Monthly SSD -- Title II'],x['Custom Recovered Monthly (Monthly Benefit)']),axis = 1)
        
                
        #Test if monthly award is higher than retro
        
        def MonthlyHigherThanRetroTester (ClosingMonthly,ClosingAward):
            ClosingMonthly = float(ClosingMonthly[1:].replace(",",""))
            ClosingAward = float(ClosingAward[1:].replace(",",""))
            
            if ClosingMonthly == 0 or ClosingAward == 0:
                return ''
            elif ClosingMonthly >= ClosingAward:
                return 'Retro Awards Should Generally Be Higher than Monthly - Confirm Amounts'
            else:
                return ''
        df['Monthly Higher than Retro Tester'] = df.apply(lambda x : MonthlyHigherThanRetroTester(x['Custom Recovered Monthly (Monthly Benefit)'],x['Custom Retro Recovery (Retroactive Award/Settlement)']),axis = 1)
        
        #Test if DAP Retro + Interim Assistance is Larger than custom Retro
        
        def DAPPlusInterimClosingTester (ClosingAward,DAPRetro,DAPInterim):
            ClosingAward = float(ClosingAward[1:].replace(",",""))
            DAPRetro = float(DAPRetro[1:].replace(",",""))
            DAPInterim = float(DAPInterim[1:].replace(",",""))
            
            if ClosingAward + 100 < DAPRetro + DAPInterim:
                return 'Closing Award Should be Larger then DAP Retro + DAP Interim'
            else:
                return ''
                
        df ['Closing Award Tester'] = df.apply(lambda x : DAPPlusInterimClosingTester(x['Custom Retro Recovery (Retroactive Award/Settlement)'],x['DAP Retro To Client'],x['DAP Interim Assistance Recovery']),axis = 1)
        
        #Education Cases should generally be $ avoided rather than awarded
        
        def EdTester (LegalProblemCode, ClosingAward):
            if LegalProblemCode.startswith("1") and ClosingAward != "$0.00":
                return "Education Cases Should have $ Avoided Rather than Awarded"
            else:
                return ""
        df['Education Award Tester'] = df.apply(lambda x : EdTester(x['Legal Problem Code'],x['Custom Retro Recovery (Retroactive Award/Settlement)']),axis = 1)
        
        #EITC tax refunds or other tax benefits, medicaid and medicare shouldn't have monthly benefits going forward
        def MonthlyTester (IOLADirectDollars,ClosingMonthly,IOLADollarSavings):
            IOLADirectDollars = str(IOLADirectDollars)
            IOLADollarSavings = str(IOLADollarSavings)
            
            if IOLADirectDollars.startswith("EITC") and ClosingMonthly != "$0.00":
                return "Tax Benefits are not Monthly Awards, please confirm"
            elif IOLADollarSavings.startswith("Health-Medic") and ClosingMonthly != "$0.00":
                return "Medicare/Medicaid Benefits are not Monthly Awards, please confirm"
            else:
                return ""
                
        df ["Monthly Tester"] = df.apply(lambda x : MonthlyTester(x['IOLA Direct Dollar Benefits to Clients'],x['Custom Recovered Monthly (Monthly Benefit)'],x['IOLA Dollar Savings to Clients']),axis = 1)
        
        
        #Food stamps and public assistance benefits don't make sense as 'avoided'
        
        def AvoidedTester (IOLADollarSavings, ClosingAvoid):
            if IOLADollarSavings == "Income Maintenance--Public Assistance" and ClosingAvoid != "$0.00":
                return "PA Benefits Should be Awards not Avoided"
            if IOLADollarSavings == "Income Maintenance-Food Stamps" and ClosingAvoid != "$0.00":
                return "SNAP Benefits Should be Awards not Avoided"
            else:
                return ''
        
        df ["Avoided Tester"] = df.apply(lambda x : AvoidedTester(x['IOLA Dollar Savings to Clients'],x['Custom Avoid (Lump Sum Avoid)']),axis = 1)
        
        #flag any retro awards that are greater than $750k
        #flag any monthly awards greater than $3k
        
        def BigAwardTester (ClosingAward,ClosingMonthly):
            ClosingMonthly = float(ClosingMonthly[1:].replace(",",""))
            ClosingAward = float(ClosingAward[1:].replace(",",""))
        
            if ClosingAward > 750000 or ClosingMonthly > 3000:
                return 'This is an unusually large award, please confirm accuracy'
            else:
                return ''
        df ["Large Award Tester"] = df.apply(lambda x : BigAwardTester(x['Custom Retro Recovery (Retroactive Award/Settlement)'],x['Custom Recovered Monthly (Monthly Benefit)']),axis = 1)
        
        #flag any dap retro awards that are greater than $100k
        #flag any monthly awards greater than $3k
        def DAPBigAwardTester (DAPAward,DAPMonthlySSI,DAPMonthlySSD):
            DAPMonthlySSI = float(DAPMonthlySSI[1:].replace(",",""))
            DAPMonthlySSD = float(DAPMonthlySSD[1:].replace(",",""))
            DAPAward = float(DAPAward[1:].replace(",",""))
        
            if DAPAward > 100000 or DAPMonthlySSI > 3000 or DAPMonthlySSD > 3000:
                return 'This is an unusually large DAP award, please confirm accuracy'
            else:
                return ''
        df ["DAP Large Award Tester"] = df.apply(lambda x : DAPBigAwardTester(x['DAP Retro To Client'],x['DAP Monthly XVI -- SSI'],x['DAP Monthly SSD -- Title II']),axis = 1)
        
        #Flag situations where a dollar amount has been added but there's no category chosen
        
        def BenefitCategoryTester(DirectDollarType, MonthlyBenefit, RetroAward):
        
            if MonthlyBenefit != "$0.00" or RetroAward != "$0.00":
                if DirectDollarType == "":
                    return "Needs Direct Dollar Benefits Type"
                else:
                    return ""
            else:
                return ""
            
        
        df["Benefit Category Tester"] = df.apply(lambda x : BenefitCategoryTester(x['IOLA Direct Dollar Benefits to Clients'],x['Custom Recovered Monthly (Monthly Benefit)'],x['Custom Retro Recovery (Retroactive Award/Settlement)']),axis =1)
        
        def SavingsCategoryTester(SavingsType, MonthlyAvoid, LumpSumAvoid):
            if MonthlyAvoid != "$0.00" or LumpSumAvoid != "$0.00":
                if SavingsType =="":
                    return "Needs Savings Type"
                else:
                    return ""
            else:
                return ""
        df["Savings Category Tester"] = df.apply(lambda x : SavingsCategoryTester(x['IOLA Dollar Savings to Clients'],x['Custom Avoid Monthly (Monthly Payment Avoided)'],x['Custom Avoid (Lump Sum Avoid)']),axis =1)
        #TesterTester 
        
        def TesterTester (DAPMonthlyTester,MonthlyHigherThanRetroTester,ClosingAwardTester,EducationAwardTester,MonthlyTester,AvoidedTester,LargeAwardTester,DAPLargeAwardTester,BenefitCategoryTester,SavingsCategoryTester):
            if DAPMonthlyTester == '' and MonthlyHigherThanRetroTester == ''and ClosingAwardTester == '' and EducationAwardTester == '' and MonthlyTester == '' and AvoidedTester == '' and LargeAwardTester == '' and DAPLargeAwardTester == '' and BenefitCategoryTester == '' and SavingsCategoryTester == '':
                return ''
            else:
                return 'Case Needs Attention'
        df ['Tester Tester'] = df.apply (lambda x : TesterTester(x['DAP Monthly $ Tester'],x['Monthly Higher than Retro Tester'],x['Closing Award Tester'],x['Education Award Tester'],x['Monthly Tester'],x['Avoided Tester'],x['Large Award Tester'],x['DAP Large Award Tester'],x['Benefit Category Tester'],x['Savings Category Tester']), axis = 1)        

        df = df[df['Tester Tester'] == "Case Needs Attention"]
        
        #Putting columns in the right order
        df = df[['Hyperlinked Case #','Assigned Branch/CC','Primary Advocate','Client First Name','Client Last Name','Date Closed','DAP Monthly $ Tester','DAP Monthly XVI -- SSI','DAP Monthly SSD -- Title II','Custom Recovered Monthly (Monthly Benefit)','Monthly Higher than Retro Tester','Custom Retro Recovery (Retroactive Award/Settlement)','Closing Award Tester','DAP Retro To Client','DAP Interim Assistance Recovery','Education Award Tester','Legal Problem Code','Monthly Tester','IOLA Direct Dollar Benefits to Clients','Benefit Category Tester','IOLA Dollar Savings to Clients','Savings Category Tester','Avoided Tester','Custom Avoid (Lump Sum Avoid)','Custom Avoid Monthly (Monthly Payment Avoided)','Large Award Tester','Custom Retro Recovery (Retroactive Award/Settlement)','DAP Large Award Tester','Outcome','Result Achieved']] 
        
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1',index=False, header = False, startrow = 1)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet.freeze_panes(1,1)
        worksheet.autofilter('A1:AI1')
        
        #create format that will make case #s look like links
        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        problem_format = workbook.add_format({
                'bg_color':'yellow',
                'text_wrap': True,
                'font_size' : 9
                })
        default_format = workbook.add_format({
                'text_wrap': True,
                'font_size' : 9
                })
        header_format = workbook.add_format({
                'text_wrap':True,
                'bold':True,
                'valign': 'middle',
                'align': 'center',
                'font_size' : 9
                })
        
        #assign new format to column A
        for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
        worksheet.set_column('A:A',15,link_format)
        worksheet.set_column('B:AI',15,default_format)
        worksheet.conditional_format('C1:BO10000',{'type': 'text',
                                             'criteria': 'containing',
                                             'value': 'Needs',
                                             'format': problem_format})        
        worksheet.conditional_format('C1:BO10000',{'type': 'text',
                                             'criteria': 'containing',
                                             'value': 'Tester',
                                             'format': problem_format})
        worksheet.conditional_format('C1:BO10000',{'type': 'text',
                                             'criteria': 'containing',
                                             'value': 'Should',
                                             'format': problem_format})
        worksheet.conditional_format('C1:BO10000',{'type': 'text',
                                             'criteria': 'containing',
                                             'value': 'please',
                                             'format': problem_format})
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Cleaned " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>IOLA $ Cleaner</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <link rel="stylesheet" href="/static/css/main.css"> 
    <h1>Flag Potential Problems with IOLA $:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Clean!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=1725" target="_blank">"Pascale Big Base Report"</a>.</li>
    <li>Once you have identified this file, click ‘Clean!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> 
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
