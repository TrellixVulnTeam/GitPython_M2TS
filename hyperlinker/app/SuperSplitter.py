#some of these imports are extraneous and left over from the flask megatutorial 
from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd
import os


@app.route("/SuperSplitter", methods=['GET', 'POST'])
def split_by_supervisor():
    #upload file from computer
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        #Test if there are extra rows at the top from a raw legalserver report - skip them if so     
        test = pd.read_excel(f)
        test.fillna('',inplace=True)
        if test.iloc[0][0] == '':
            data_xls = pd.read_excel(f,skiprows=2)
        else:
            data_xls = pd.read_excel(f)
        
        if 'Primary Assignment' not in data_xls.columns:
            data_xls['Primary Assignment'] = data_xls['Primary Advocate']
        
        if 'Matter/Case ID#' not in data_xls.columns:
            data_xls['Matter/Case ID#'] = data_xls['Hyperlinked Case #']
        
        del data_xls['Hyperlinked Case #']
        
        #apply hyperlink methodology with splicing and concatenation
      
        def NoIDDelete(CaseID):
            if CaseID == '':
                return 'No Case ID'
            else:
                return str(CaseID)
        data_xls['Matter/Case ID#'] = data_xls.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = data_xls['Matter/Case ID#']
        data_xls['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del data_xls['Matter/Case ID#']
        move=data_xls['Temp Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        JoseAbrigo_Supervisees = [
        "Bromberg, Iris",
        "Torres, Jasmin",
        "Granfield, Rachel",
        "Anderson, Theresa"
        ]
        
        ShantonuBasu_Supervisees = [
        "Shah, Ami",
        "Sharma, Sagar",
        "Abbas, Sayeda",
        "Evers, Erin",
        "Spencer, Eleanor"
        ]
        
        EricaBraudy_Supervisees = [
        "Mercedes, Jannelys",
        "Kulig, Jessica",
        "James, Lelia",
        "Sanchez, Dennis"
        ]
        
        ArchanaDittakavi_Supervisees = [
        "Briggs, John",
        "Mottley, Darlene",
        "Porcelli, Ronald",
        "Gonzalez-Munoz, Rossana",
        "Honan, Thomas",
        "Kelly, Kitanya",
        "Almanzar, Milagros",
        "Yamasaki, Emily Woo"
        ]
        
        DeniseAcron_Supervisees = [
        "Anunkor, Ifeoma",
        "Mendia-Yadaicela, Michelle",
        "Reyes, Nicole"
        ]
        
        RosalindBlack_Supervisees = [
        "Dittakavi, Archana",
        "Basu, Shantonu",
        "Freeman, Daniel",
        "Braudy, Erica",
        "Frierson, Jerome",
        "Heller, Steven"
        ]
        
        TanyaDouglas_Supervisees = [
        "Kaplan, William",
        "Sun, Dao",
        "Williams, Gerald",
        "Robles, Eliana",
        "Heaton, Betty",
        "Isaacs, Nicole"
        ]
        
        PeggyEarisman_Supervisees = [
        "Martinez-Gunter, Maribel",
        "Black, Rosalind",
        "Douglas, Tanya",
        "Goldman, Caitlin",
        "Abrigo, Jose"
        ]
        
        DanielFreeman_Supervisees = [
        "Labossiere, Samantha",
        "Gonzalez, Matias",
        "Gokhale, Aparna",
        "Vega, Rita"
        ]
        
        JeromeFrierson_Supervisees = [
        "Duffy-Greaves, Kevin",
        "Ortiz, Matthew",
        "Allen, Sharette"
        ]
        
        CaitlinGoldman_Supervisees = [
        "Pepitone, Dan",
        "McCune, Mary",
        "Rosner, Julia",
        "Martinez Alonzo, Washcarina",
        "Brito, Victor"
        ]
        
        StevenHeller_Supervisees = [
        "Latterner, Matt",
        "Delgadillo, Omar",
        "Robles-Castillo, Camila",
        "Almanzar, Yocari"
        ]
        
        LuisHenriquez_Supervisees = [
        "Acron, Denise"
        ]
        
        MaribelMartinezGunter_Supervisees = [
        "Trinidad, Lenina",
        "Caban-Gandhi, Celina",
        "Restrepo-Serrano, Francois",
        "Singh, Ermela",
        "Carlier, Milton",
        "Ventura, Lynn",
        "Patel, Roopal",
        "Guerra, Yolanda",
        "Sanchez, Ingrid",
        "Smith-Menjivar, Eleise",
        "Maldonando, Brenda"
        ]
        
        AmiShah_Supervisees = [
        "Wilkes, Nicole",
        "He, Ricky",
        "Hao, Lindsay"
        ]
        
        def SuperAssign(advocate):
            if advocate in JoseAbrigo_Supervisees:
                return "Jose Abrigo"
            elif advocate in ShantonuBasu_Supervisees:
                return "Shantonu Basu"
            elif advocate in EricaBraudy_Supervisees:
                return "Erica Braudy"
            elif advocate in ArchanaDittakavi_Supervisees:
                return "Archana Dittakavi"
            elif advocate in DeniseAcron_Supervisees:
                return "Denise Acron"
            elif advocate in RosalindBlack_Supervisees:
                return "Rosalind Black"
            elif advocate in TanyaDouglas_Supervisees:
                return "Tanya Douglas"
            elif advocate in PeggyEarisman_Supervisees:
                return "Peggy Earisman"
            elif advocate in DanielFreeman_Supervisees:
                return "Daniel Freeman"
            elif advocate in JeromeFrierson_Supervisees:
                return "Jerome Frierson"
            elif advocate in CaitlinGoldman_Supervisees:
                return "Caitlin Goldman"
            elif advocate in StevenHeller_Supervisees:
                return "Steven Heller"
            elif advocate in LuisHenriquez_Supervisees:
                return "Luis Henriquez"
            elif advocate in MaribelMartinezGunter_Supervisees:
                return "Maribel Martinez-Gunter"
            elif advocate in AmiShah_Supervisees:
                return "Ami Shah"
            else:
                return "Unaffiliated"
        
        data_xls["Supervisor"] = data_xls.apply(lambda x: SuperAssign(x['Primary Assignment']), axis = 1)
        
        data_xls = data_xls.sort_values(by=['Primary Assignment'])
        
        #split into separate dataframes
        
        Jose_data_xls = data_xls[data_xls['Supervisor'] == 'Jose Abrigo']
        Shantonu_data_xls = data_xls[data_xls['Supervisor'] == 'Shantonu Basu']
        Erica_data_xls = data_xls[data_xls['Supervisor'] == 'Erica Braudy']
        Archana_data_xls = data_xls[data_xls['Supervisor'] == 'Archana Dittakavi']
        Denise_data_xls = data_xls[data_xls['Supervisor'] == 'Denise Acron']
        Rosalind_data_xls = data_xls[data_xls['Supervisor'] == 'Rosalind Black']
        Tanya_data_xls = data_xls[data_xls['Supervisor'] == 'Tanya Douglas']
        Peggy_data_xls = data_xls[data_xls['Supervisor'] == 'Peggy Earisman']
        Daniel_data_xls = data_xls[data_xls['Supervisor'] == 'Daniel Freeman']
        Jerome_data_xls = data_xls[data_xls['Supervisor'] == 'Jerome Frierson']
        Caitlin_data_xls = data_xls[data_xls['Supervisor'] == 'Caitlin Goldman']
        Steven_data_xls = data_xls[data_xls['Supervisor'] == 'Steven Heller']
        Luis_data_xls = data_xls[data_xls['Supervisor'] == 'Luis Henriquez']
        Maribel_data_xls = data_xls[data_xls['Supervisor'] == 'Maribel Martinez-Gunter']
        Ami_data_xls = data_xls[data_xls['Supervisor'] == 'Ami Shah']
        Unaffiliated_data_xls = data_xls[data_xls['Supervisor'] == 'Unaffiliated']
        
        #bounce worksheets back to excel
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        
        Jose_data_xls.to_excel(writer, sheet_name='JoseAbrigo',index=False)
        Shantonu_data_xls.to_excel(writer, sheet_name='ShantonuBasu',index=False)
        Erica_data_xls.to_excel(writer, sheet_name='EricaBraudy',index=False)
        Archana_data_xls.to_excel(writer, sheet_name='ArchanaDittakavi',index=False)
        Denise_data_xls.to_excel(writer, sheet_name='DeniseAcron',index=False)
        Rosalind_data_xls.to_excel(writer, sheet_name='RosalindBlack',index=False)
        Tanya_data_xls.to_excel(writer, sheet_name='TanyaDouglas',index=False)
        Peggy_data_xls.to_excel(writer, sheet_name='PeggyEarisman',index=False)
        Daniel_data_xls.to_excel(writer, sheet_name='DanielFreeman',index=False)
        Jerome_data_xls.to_excel(writer, sheet_name='JeromeFrierson',index=False)
        Caitlin_data_xls.to_excel(writer, sheet_name='CaitlinGoldman',index=False)
        Steven_data_xls.to_excel(writer, sheet_name='StevenHeller',index=False)
        Luis_data_xls.to_excel(writer, sheet_name='LuisHenriquez',index=False)
        Maribel_data_xls.to_excel(writer, sheet_name='MaribelMartinezGunter',index=False)
        Ami_data_xls.to_excel(writer, sheet_name='AmiShah',index=False)
        Unaffiliated_data_xls.to_excel(writer, sheet_name='Unaffiliated',index=False)
        
        
        workbook = writer.book
        
        worksheetJose = writer.sheets['JoseAbrigo']
        worksheetShantonu = writer.sheets['ShantonuBasu']
        worksheetErica = writer.sheets['EricaBraudy']
        worksheetArchana = writer.sheets['ArchanaDittakavi']
        worksheetDenise = writer.sheets['DeniseAcron']
        worksheetRosalind = writer.sheets['RosalindBlack']
        worksheetTanya = writer.sheets['TanyaDouglas']
        worksheetPeggy = writer.sheets['PeggyEarisman']
        worksheetDaniel = writer.sheets['DanielFreeman']
        worksheetJerome = writer.sheets['JeromeFrierson']
        worksheetCaitlin = writer.sheets['CaitlinGoldman']
        worksheetSteven = writer.sheets['StevenHeller']
        worksheetLuis = writer.sheets['LuisHenriquez']
        worksheetMaribel = writer.sheets['MaribelMartinezGunter']
        worksheetAmi = writer.sheets['AmiShah']
        worksheetUnaffiliated = writer.sheets['Unaffiliated']


        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        
        worksheetJose.freeze_panes(1, 1)
        worksheetShantonu.freeze_panes(1, 1)
        worksheetErica.freeze_panes(1, 1)
        worksheetArchana.freeze_panes(1, 1)
        worksheetDenise.freeze_panes(1, 1)
        worksheetRosalind.freeze_panes(1, 1)
        worksheetTanya.freeze_panes(1, 1)
        worksheetPeggy.freeze_panes(1, 1)
        worksheetDaniel.freeze_panes(1, 1)
        worksheetJerome.freeze_panes(1, 1)
        worksheetCaitlin.freeze_panes(1, 1)
        worksheetSteven.freeze_panes(1, 1)
        worksheetLuis.freeze_panes(1, 1)
        worksheetMaribel.freeze_panes(1, 1)
        worksheetAmi.freeze_panes(1, 1)
        worksheetUnaffiliated.freeze_panes(1, 1)
        
        worksheetJose.set_column('A:A',20,link_format)
        worksheetShantonu.set_column('A:A',20,link_format)
        worksheetErica.set_column('A:A',20,link_format)
        worksheetArchana.set_column('A:A',20,link_format)
        worksheetDenise.set_column('A:A',20,link_format)
        worksheetRosalind.set_column('A:A',20,link_format)
        worksheetTanya.set_column('A:A',20,link_format)
        worksheetPeggy.set_column('A:A',20,link_format)
        worksheetDaniel.set_column('A:A',20,link_format)
        worksheetJerome.set_column('A:A',20,link_format)
        worksheetCaitlin.set_column('A:A',20,link_format)
        worksheetSteven.set_column('A:A',20,link_format)
        worksheetLuis.set_column('A:A',20,link_format)
        worksheetMaribel.set_column('A:A',20,link_format)
        worksheetAmi.set_column('A:A',20,link_format)
        worksheetUnaffiliated.set_column('A:A',20,link_format)
        
        worksheetJose.set_column('B:ZZ',25)
        worksheetShantonu.set_column('B:ZZ',25)
        worksheetErica.set_column('B:ZZ',25)
        worksheetArchana.set_column('B:ZZ',25)
        worksheetDenise.set_column('B:ZZ',25)
        worksheetRosalind.set_column('B:ZZ',25)
        worksheetTanya.set_column('B:ZZ',25)
        worksheetPeggy.set_column('B:ZZ',25)
        worksheetDaniel.set_column('B:ZZ',25)
        worksheetJerome.set_column('B:ZZ',25)
        worksheetCaitlin.set_column('B:ZZ',25)
        worksheetSteven.set_column('B:ZZ',25)
        worksheetLuis.set_column('B:ZZ',25)
        worksheetMaribel.set_column('B:ZZ',25)
        worksheetAmi.set_column('B:ZZ',25)
        worksheetUnaffiliated.set_column('B:ZZ',25)
        
        writer.save()
        
        #send file back to user
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "SuperSplit " + f.filename)

        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>Supervisor Splitter</title>
    <link rel="stylesheet" href="/static/css/main.css">  
    <h1>Split Your Spreadsheet by Supervisor:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=Split!>
    </form>
    
    
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to split into different documents by supervisor.</li> 
    <li>Once you have identified this file, click ‘Split!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li>
    <li>Note, the column with your case ID numbers in it must be titled "Matter/Case ID#" or "id" for this to work.</li>
    </ul>
    </br>
    <a href="/">Home</a>
    '''
    
