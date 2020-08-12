from flask import request, send_from_directory
from app import app, DataWizardTools
import pandas as pd
import datetime
import numpy

@app.route("/PTOEstimator", methods=['GET', 'POST'])
def PTOEstimator():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        df = pd.read_excel(f,skiprows=2)
        df.fillna('',inplace=True)
        
        
        def RemoveNoPayPeriod(CaseID):
            CaseID=str(CaseID)
            if CaseID == '' or CaseID == 'nan' or CaseID.startswith("Unique") == True:
                return 'No Case ID'

            else:
                return str(CaseID)
        
        df['Pay Period'] = df.apply(lambda x : DataWizardTools.RemoveNoCaseID(x['Pay Period']),axis=1)        
        df = df[df['Pay Period'] != 'No Case ID']
        
        def DateMaker (Date):
            
            if Date == "":
                return ""
            else:
                DateMonth = Date[:2]
                DateYear = Date[6:]
                return int(DateYear + DateMonth)


        df['YearMonth'] = df.apply(lambda x: DateMaker(x['Date of Service']), axis=1)
        
        #add annual and personal together
        
        def PTOCombiner(ActivityType,TimeSpent):
            if ActivityType == "Annual" or ActivityType == "Personal":
                return TimeSpent
            else:
                return 0
        
        df['PTO Expended'] = df.apply(lambda x: PTOCombiner(x['Activity Type'], x['Time Spent']), axis=1)
        
        #Construct Summary Tables
        PTO_pivot = pd.pivot_table(df,index=['YearMonth'],values=['PTO Expended'],aggfunc=sum,fill_value='blank')

        #How much PTO have you accumulated
        
        FirstTwelve = PTO_pivot.head(n=12)
        
        def PTOAccumulater(YearMonth):
            
            Month = str(YearMonth)[-2:]
            if YearMonth in FirstTwelve.index and Month == "01":
                return 7
            elif YearMonth in FirstTwelve.index:
                return 14
            elif Month == "01" or Month == "04" or Month == "07" or Month == "10":
                return 21
            elif Month == "02" or Month == "03" or Month == "05" or Month == "06" or Month == "08" or Month == "09" or Month == "11" or Month == "12":
                return 14

        PTO_pivot['PTO Accumulated'] = PTO_pivot.apply(lambda x: PTOAccumulater(x.name), axis = 1)
        PTO_pivot['PTO'] = 'PTO'

        
        print(FirstTwelve)

        #Topline Summary

        PTO_totals = pd.pivot_table(PTO_pivot,index=['PTO'],values=['PTO Expended','PTO Accumulated'],aggfunc=sum,fill_value='blank')


        #put sheets in right order
        df = df[['Pay Period','Date of Service','YearMonth','Time Spent','Activity Type','PTO Expended']]
        

        #Bounce to a new excel file        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        PTO_totals.to_excel(writer, sheet_name='PTO Totals')
        PTO_pivot.to_excel(writer, sheet_name='PTO Summary')
        df.to_excel(writer, sheet_name='Time Entry List',index=False)
        
        

        workbook = writer.book
        PTOTotals = writer.sheets['PTO Totals']
        PTOSummary = writer.sheets['PTO Summary']
        TimeEntryList = writer.sheets['Time Entry List']
        
        PTOTotals.set_column('A:D',20)
        TimeEntryList.set_column('A:I',20)
        PTOSummary.set_column('A:F',20)

        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Estimated " + f.filename)

    return '''
    <!doctype html>
    <title>PTO Estimator</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Estimate How Much PTO You Have Expended and Accrued:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Estimate!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>Here's a link to the LegalServer report this works with: <a href="https://lsnyc.legalserver.org/report/dynamic?load=505" target="_blank">"Administrative Leave Time"</a></li>
    <li>Run this report for your own name and with the pay period filter set from 0 to 5000</li>
    <li>This should work for anybody who started since 2012 and has been consistenly employed full-time since then.</li>
    <li>I do not vouch for accuracy but this is my best bet! </li>
    </ul>
    <a href="/">Home</a>
    '''
