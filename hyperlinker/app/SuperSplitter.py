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
            df = pd.read_excel(f,skiprows=2)
        else:
            df = pd.read_excel(f)
        
        
        if 'Matter/Case ID#' not in df.columns:
            try:
                df['Matter/Case ID#'] = df['Hyperlinked Case #']
                del df['Hyperlinked Case #']
            except: 
                df['Matter/Case ID#'] = df['Hyperlinked CaseID#']
                del df['Hyperlinked CaseID#']
        
        
        #apply hyperlink methodology with splicing and concatenation
      
        def NoIDDelete(CaseID):
            if CaseID == '':
                return 'No Case ID'
            else:
                return str(CaseID)
        df['Matter/Case ID#'] = df.apply(lambda x: NoIDDelete(x['Matter/Case ID#']), axis=1)
        
        last7 = df['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = df['Matter/Case ID#']
        df['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        del df['Matter/Case ID#']
        move=df['Temp Hyperlinked Case #']
        df.insert(0,'Hyperlinked Case #', move)           
        del df['Temp Hyperlinked Case #']
        
        JoseAbrigo_Supervisees = [
        "Torres, Jasmin",
        "Torres, Jasmin E",
        "Granfield, Rachel",
        "Granfield, Rachel W",
        "Anderson, Theresa"
        ]
        
        ShantonuBasu_Supervisees = [
        "Shah, Ami",
        "Shah, Ami Mahendra",
        "Sharma, Sagar",
        "Grater, Ashley",
        "Arboleda, Heather",
        "Evers, Erin",
        "Evers, Erin C.",
        "Spencer, Eleanor",
        "Spencer, Eleanor G"
        ]
        
        EricaBraudy_Supervisees = [
        "Mercedes, Jannelys",
        "Mercedes, Jannelys J",
        "Sanchez, Dennis",
        "Gelly-Rahim, Jibril",
        "Flores, Aida",
        "Harshberger, Sae"

        ]
        
        ArchanaDittakavi_Supervisees = [
        "Briggs, John",
        "Briggs, John M",
        "Gonzalez-Munoz, Rossana",
        "Gonzalez-Munoz, Rossana G",
        "Honan, Thomas",
        "Honan, Thomas J",
        "Yamasaki, Emily Woo",
        "Yamasaki, Emily Woo J"
        ]
        
        DeniseAcron_Supervisees = [
        "Anunkor, Ifeoma",
        "Mendia-Yadaicela, Michelle",
        "Reyes, Nicole",
        "Vega, Rita"

        ]
        
        RosalindBlack_Supervisees = [
        "Dittakavi, Archana",
        "Basu, Shantonu",
        "Basu, Shantonu J",
        "Freeman, Daniel",
        "Braudy, Erica",
        "Frierson, Jerome",
        "Frierson, Jerome C",
        "Heller, Steven",
        "Heller, Steven E",
        "Sun, Dao"
        ]
        
        TanyaDouglas_Supervisees = [
        "Kaplan, William",
        "Kaplan, William D.",
        "Williams, Gerald",
        "Williams, Gerald S",
        "Robles, Eliana",
        "Isaacs, Nicole"
        ]
        
        PeggyEarisman_Supervisees = [
        "Martinez-Gunter, Maribel",
        "Black, Rosalind",
        "Douglas, Tanya",
        "Douglas, Tanya M.",
        "Goldman, Caitlin",
        "Goldman, Caitlin S",
        "Abrigo, Jose",
        "Abrigo, Jose I",
        "Henriquez, Luis A"
        ]
        
        DaoSun_Supervisees = [
        "Agarwal, Rachna",
        "Gonzalez, Matias G",
        "Gonzalez, Matias",
        "Gokhale, Aparna S"
        ]
        

        JeromeFrierson_Supervisees = [
        "Duffy-Greaves, Kevin",
        "Saxton, Jonathan G",
        "Allen, Sharette",
        "Orsini, Mary",
        "Harris, Tycel"
        ]
        
        CaitlinGoldman_Supervisees = [
        "McCune, Mary",
        "Rosner, Julia",
        "Rosner, Julia P",
        "Martinez Alonzo, Washcarina",
        "Martinez Alonzo, Washcarina B",
        "Brito, Victor"
        ]
        
        StevenHeller_Supervisees = [
        "Latterner, Matt",
        "Latterner, Matt J",
        "Delgadillo, Omar",
        "Avila, Giselle",
        "Almanzar, Yocari"
        ]
        
        LuisHenriquez_Supervisees = [
        "Acron, Denise"
        ]
        
        MaribelMartinezGunter_Supervisees = [
        "Trinidad, Lenina",
        "Trinidad, Lenina C.",
        "Caban-Gandhi, Celina",
        "Restrepo-Serrano, Francois",
        "Restrepo-Serrano, Francois M",
        "Guerra, Yolanda",
        "Sanchez, Ingrid",
        "Smith-Menjivar, Eleise",
        "Smith-Menjivar, Eleise C",
        "Maldonando, Brenda",
        "Sambartaro, Debra"
        ]
        
        YolandaGuerra_Supervisees = [
        "Singh, Ermela",
        "Carlier, Milton",
        "Ventura, Lynn",
        "Patel, Roopal",
        "Patel, Roopal B",
        ]
        
        AmiShah_Supervisees = [
        "Wilkes, Nicole",
        "He, Ricky",
        "Hao, Lindsay",
        "Abbas, Sayeda",
        "Risener, Jennifer",
        "Surface, Ben"
        ]
        
        SayedaAbbas_Supervisees = [
        "Bocangel, Maricella",
        "Chow, Corrin"
        ]
        
        ThomasHonan_Supervisees = [
        "James, Lelia",
        "Kelly, Kitanya",
        "Mattar, Amira",
        "Whedon, Rebecca V",
        "Almanzar, Milagros"
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
            elif advocate in DaoSun_Supervisees:
                return "Dao Sun"
            elif advocate in ThomasHonan_Supervisees:
                return "Thomas Honan"
            elif advocate in YolandaGuerra_Supervisees:
                return "Yolanda Guerra"                
            else:
                return "zzNo Assignment"
        
        df["Supervisor"] = df.apply(lambda x: SuperAssign(x['Primary Advocate Name']), axis = 1)
        
        df = df.sort_values(by=['Primary Advocate Name'])
        
        #split into separate dataframes
        df_dictionary = dict(tuple(df.groupby('Supervisor')))
        
        #bounce worksheets back to excel
        def save_xls(dict_df, path):
            writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
            for i in dict_df:
                dict_df[i].to_excel(writer, i, index = False)
                workbook = writer.book
                ws = writer.sheets[i]
                link_format = workbook.add_format({'font_color':'blue','bold':True,'underline':True})
                ws.set_column('A:A',20,link_format)
                ws.set_column('B:ZZ',25)
                ws.autofilter('B1:ZZ1')
                ws.freeze_panes(1, 2)

            writer.save()
        
        output_filename = f.filename
        
        save_xls(dict_df = df_dictionary, path = "app\\sheets\\" + output_filename)
   
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Split " + f.filename)
        
#what the user-facing site looks like
    return '''
    <!doctype html>
    <title>MLS Supervisor Splitter</title>
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
    
