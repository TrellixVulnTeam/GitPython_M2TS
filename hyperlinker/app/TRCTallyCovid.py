from flask import request, send_from_directory
from app import app, DataWizardTools, HousingToolBox
import pandas as pd
from datetime import date
import numpy

@app.route("/TRCtallyCovid", methods=['GET', 'POST'])
def upload_TRCtallyCovid():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        data_xls = pd.read_excel(f)
        data_xls.fillna('',inplace=True)
        data_xls['city'] = data_xls.apply(lambda x: DataWizardTools.QueensConsolidater(x['city']), axis = 1)
        data_xls['city'] = data_xls['city'].str.upper()
        data_xls['city'] = data_xls['city'].str.replace('NEW YORK','MANHATTAN')
        
        
        today = date.today()
        print(today.month)
        
        if today.month >= 8:
            howmanymonths = today.month - 7
        else:
            howmanymonths = today.month + 12 - 7
        
        print(howmanymonths)
        
        data_xls['EligibilityDateConstruct'] = data_xls.apply(lambda x: DataWizardTools.DateMaker(x['eligibility_date']), axis=1)
        
        data_xls['Pre-3/1/20 Elig Date?'] = data_xls.apply(lambda x: HousingToolBox.PreThreeOne(x['EligibilityDateConstruct']), axis=1)
        
        #Value of Case
                     
        def CaseValue(service_type,waiver,referral_source,caseid,PreThreeOne):
            if service_type == 'Full Rep' or service_type == 'Pre-Litigation Strategies' or service_type == 'Representation—EOIR':
                return 1
            elif service_type.startswith("Advice Only") == True and waiver == 'FJC Waiver':
                return 1
            elif service_type.startswith("Advice Only") == True and referral_source == 'Family Justice Center':
                return 1
            elif service_type.startswith("Advice Only") == True and PreThreeOne == "Yes":
                return .25
            elif service_type.startswith("Advice Only") == True and PreThreeOne == "No":
                return 1
            elif caseid == '':
                return ''
            else:
                return 0
        
        data_xls['Case Value'] = data_xls.apply(lambda x: CaseValue(x['service_type'], x['waiver'],x['referral_source'],x['id'],x['Pre-3/1/20 Elig Date?']), axis=1)
        
        
        
        #Assign Zips to Deliverable Categories
        
        def ServiceArea (zip,city):
            if zip == 10453 or zip == 10452:
                return "Bronx - Morris Height/Highbridge"
            elif zip == 10459 or zip == 10457 or zip == 10460:
                return "Bronx - Longwood/East Tremont/West Farms"
            elif zip == 11206 or zip == 11237:
                return "Brooklyn - Ridgewood/Bushwick"
            elif zip == 11215 or zip == 11217 or zip == 11231:
                return "Brooklyn - Gowanus/Park Slope/Boerum Hill/Carroll Garden/Red Hook"
            elif zip == 11207 or zip == 11208 or zip == 11212 or zip == 11233:
                return "Brooklyn - East New York/Brownsville/Ocean Hill"
            
            elif zip == 10029 or zip == 10035:
                return "Manhattan - East Harlem"
            elif zip == 10034:
                return "Manhattan - Inwood"
            elif zip == 10032 or zip == 10033:
                return "Manhattan - Washington Heights"
            elif zip == 11101:
                return "Queens - Long Island City"
            elif zip == 11354 or zip == 11358:
                return "Queens - Flushing/West Flushing"
            elif zip == 11691 or zip == 11692:
                return "Queens - Far Rockaway"
            elif zip == 10304 or zip == 10301:
                return "Staten Island - Stapleton/Bay Street"
            elif city == "MANHATTAN":
                return "Manhattan - Other Zips"
            elif city == "BROOKLYN":
                return "Brooklyn - Other Zips"
            elif city == "BRONX":
                return "Bronx - Other Zips"
            elif city == "QUEENS":
                return "Queens - Other Zips"
            elif city == "STATEN ISLAND":
                return "Staten Island - Other Zips"
            else:  
                return ""
        
        data_xls['Service Area'] = data_xls.apply(lambda x: ServiceArea(x['zip'], x['city']), axis = 1)
        
        #pulling month for later
        data_xls['Eligibility Month'] = pd.to_numeric(data_xls['eligibility_date'].apply(lambda x: str(x)[:2]))
        

        
        #extra case value column for tallying later and not getting problems with duplicate-named columns
        data_xls['Case Value Sum'] = data_xls['Case Value']
        
      
        
        #Construct Summary Tables
        city_pivot = pd.pivot_table(data_xls,index=['city'],values=['Case Value'],aggfunc=sum,fill_value=0)
        
        city_reorder = ["BRONX","BROOKLYN","MANHATTAN","QUEENS","STATEN ISLAND"]
        
        city_pivot = city_pivot.reindex(city_reorder)
        
        area_pivot = pd.pivot_table(data_xls,index=['Service Area'],values=['Case Value'],aggfunc=sum,fill_value=0)
        
        area_reorder = ["Bronx - Morris Height/Highbridge","Bronx - Longwood/East Tremont/West Farms","Bronx - Other Zips","Brooklyn - Ridgewood/Bushwick","Brooklyn - Gowanus/Park Slope/Boerum Hill/Carroll Garden/Red Hook","Brooklyn - East New York/Brownsville/Ocean Hill","Brooklyn - Other Zips","Manhattan - East Harlem","Manhattan - Inwood","Manhattan - Washington Heights","Manhattan - Other Zips","Queens - Long Island City","Queens - Flushing/West Flushing","Queens - Far Rockaway","Queens - Other Zips","Staten Island - Stapleton/Bay Street","Staten Island - Other Zips"]
        
        area_pivot = area_pivot.reindex(area_reorder)
        
        zip_pivot = pd.pivot_table(data_xls,index=['zip'],values=['Case Value'],aggfunc=sum,fill_value=0)
        
        advice_proportion = pd.pivot_table(data_xls,index=['city','Case Value'],values=['Case Value Sum'],aggfunc=sum,fill_value=0)
        
        dummy_advice_proportion = pd.pivot_table(data_xls,index=['city'],values=['Case Value Sum'],aggfunc=sum,fill_value=0)
        
        
        #Add Goals to Summary Tables:
                
        def BoroughGoal(city):
            if city == "BRONX":
                return 1700
            elif city == "BROOKLYN":
                return 2631
            elif city == "MANHATTAN":
                return 1316
            elif city == "QUEENS":
                return 379
            elif city == "STATEN ISLAND":
                return 329
            elif city == "":
                return ""
            else:
                return "hmmmm!"
        
        city_pivot.reset_index(inplace=True)
        
               
        city_pivot['Annual Goal'] = city_pivot.apply(lambda x: BoroughGoal(x['city']), axis=1)
        
        city_pivot['Proportional Goal'] = numpy.ceil(city_pivot['Annual Goal']/12) * howmanymonths
        
  
        
        def AreaGoal(area):
            if area == "Bronx - Morris Height/Highbridge":
                return 1100
            elif area == "Bronx - Longwood/East Tremont/West Farms":
                return 260
            elif area == "Bronx - Other Zips":
                return 340
            elif area == "Brooklyn - Ridgewood/Bushwick":
                return 69
            elif area == "Brooklyn - Gowanus/Park Slope/Boerum Hill/Carroll Garden/Red Hook":
                return 36
            elif area == "Brooklyn - East New York/Brownsville/Ocean Hill":
                return 1650
            elif area == "Brooklyn - Other Zips":
                return 876
            
            elif area == "Manhattan - East Harlem":
                return 618
            elif area == "Manhattan - Inwood":
                return 238
            elif area == "Manhattan - Washington Heights":
                return 100
            elif area == "Manhattan - Other Zips":
                return 360
            elif area == "Queens - Long Island City":
                return 40
            elif area == "Queens - Flushing/West Flushing":
                return 72
            elif area == "Queens - Far Rockaway":
                return 156
            elif area == "Queens - Other Zips":
                return 111
            elif area == "Staten Island - Stapleton/Bay Street":
                return 263
            elif area == "Staten Island - Other Zips":
                return 66
            elif area == "":
                return ""
            else:
                return "hmmmmmmmm!"
                
        area_pivot.reset_index(inplace=True)
        
        area_pivot['Annual Goal'] = area_pivot.apply(lambda x: AreaGoal(x['Service Area']), axis=1)
        
        #make it so that it's tallying goals on a cumulative basis
        
       
        area_pivot['Proportional Goal'] = numpy.ceil(area_pivot['Annual Goal']/12)*howmanymonths
        
        
        #Add percentage calculator:
        city_pivot['Proportional Percentage']=city_pivot['Case Value']/city_pivot['Proportional Goal']
        city_pivot['Annual Percentage']=city_pivot['Case Value']/city_pivot['Annual Goal']
        area_pivot['Annual Percentage']=area_pivot['Case Value']/area_pivot['Annual Goal']
        area_pivot['Proportional Percentage']=area_pivot['Case Value']/area_pivot['Proportional Goal']
        advice_proportion.at['BRONX','city total enrollments'] = dummy_advice_proportion.at['BRONX','Case Value Sum']
        advice_proportion.at['BROOKLYN','city total enrollments'] = dummy_advice_proportion.at['BROOKLYN','Case Value Sum']
        advice_proportion.at['MANHATTAN','city total enrollments'] = dummy_advice_proportion.at['MANHATTAN','Case Value Sum']
        advice_proportion.at['QUEENS','city total enrollments'] = dummy_advice_proportion.at['QUEENS','Case Value Sum']
            #specific to advice proportion calculations
        advice_proportion.at['STATEN ISLAND','city total enrollments'] = dummy_advice_proportion.at['STATEN ISLAND','Case Value Sum']
        advice_proportion['Percentage']=advice_proportion['Case Value Sum']/advice_proportion['city total enrollments']
        del advice_proportion['city total enrollments']
        #advice_proportion = advice_proportion.drop('')
        
        
        #put sheets in right order
        
        data_xls = data_xls[['id','city','street_number','Street','zip','Service Area','service_type','waiver','referral_source','Eligibility Month','Pre-3/1/20 Elig Date?','Case Value']]
        city_pivot = city_pivot[['city','Case Value','Proportional Goal','Proportional Percentage','Annual Goal','Annual Percentage']]
        area_pivot = area_pivot[['Service Area','Case Value','Proportional Goal','Proportional Percentage','Annual Goal','Annual Percentage']]
        
        
        #Bounce to a new excel file        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Case List',index=False)
        city_pivot.to_excel(writer, sheet_name='City Pivot',index=False)
        area_pivot.to_excel(writer, sheet_name='Service Area Pivot',index=False)
        zip_pivot.to_excel(writer, sheet_name='Zip Code Pivot')
        advice_proportion.to_excel(writer, sheet_name='Proportion of Advice Cases')
        #dummy_advice_proportion.to_excel(writer, sheet_name='Dummy Pro')
        
        workbook = writer.book
        CaseList = writer.sheets['Case List']
        CityPivot = writer.sheets['City Pivot']
        AreaPivot = writer.sheets ['Service Area Pivot']
        ZipPivot = writer.sheets ['Zip Code Pivot']
        ProportionTable = writer.sheets ['Proportion of Advice Cases']
        
        percent_format = workbook.add_format()
        percent_format.set_num_format('0.00%')
        
        CaseList.set_column('A:I',20)
        CityPivot.set_column('A:F',20)
        CityPivot.set_column('D:D',23,percent_format)
        CityPivot.set_column('F:F',20,percent_format)
        AreaPivot.set_column('A:A', 65)
        AreaPivot.set_column('B:F', 20)
        AreaPivot.set_column('D:D',23,percent_format)
        AreaPivot.set_column('F:F',20,percent_format)
        ZipPivot.set_column('A:B',20)
        ProportionTable.set_column('A:D',15)
        ProportionTable.set_column('D:D',20,percent_format)
        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Tallied " + f.filename)

    return '''
    <!doctype html>
    <title>TRC Tally [Covid]</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Tally your TRC Cases against Goals:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=TRC-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with housing reports that have been processed by Kim Robinson.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for TRC.</li> 
    <li>Once you have identified this file, click ‘TRC-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
