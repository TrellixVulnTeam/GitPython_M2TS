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

        def DateMaker (Date):
            
            if Date == "":
                return ""
            else:
                DateMonth = Date[:2]
                DateYear = Date[6:]
                return int(DateYear + DateMonth)


        df['YearMonth'] = df.apply(lambda x: DateMaker(x['Date of Service']), axis=1)
        
        
        #Construct Summary Tables
        PTO_pivot = pd.pivot_table(df[df['Activity Type'] == 'Annual'],index=['Activity Type','YearMonth'],values=['Time Spent'],aggfunc=sum,fill_value='Time Spent')



        #put sheets in right order
        df = df[['Pay Period','Date of Service','YearMonth','Time Spent','Activity Type']]


        #Bounce to a new excel file        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name='Time Entry List',index=False)
        PTO_pivot.to_excel(writer, sheet_name='PTO Summary')
        

        workbook = writer.book
        PTOSummary = writer.sheets['PTO Summary']
        TimeEntryList = writer.sheets['Time Entry List']
        

        TimeEntryList.set_column('A:I',20)
        PTOSummary.set_column('A:F',20)

        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Estimated " + f.filename)

    return '''
    <!doctype html>
    <title>PTO Estimator</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Estimate How Much PTO You Have Accrued:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Estimate!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>I do not vouch for accuracy but this is my best bet! </ul>
    </br>
    <a href="/">Home</a>
    '''
