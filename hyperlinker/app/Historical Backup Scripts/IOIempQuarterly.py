from flask import render_template, flash, redirect, url_for, request, Flask, jsonify, send_from_directory
from app import app, db
from app.models import User, Post
from app.forms import PostForm
from werkzeug.urls import url_parse
from datetime import datetime
import pandas as pd

#list of previously reported cases for 'substantial activity' question
ReportedFY19= ["All Cases Reported FY19",
                "19-1896196",
                "18-1880946",
                "19-1893307",
                "19-1889209",
                "19-1899655",
                "18-1875345",
                "19-1903550",
                "18-1880465",
                "19-1889319",
                "19-1890134",
                "19-1902591",
                "19-1900426",
                "19-1899943",
                "19-1897491",
                "18-1884254",
                "18-1881815",
                "18-1881138",
                "18-1879579",
                "19-1897882",
                "18-1874750",
                "18-1885541",
                "19-1891574",
                "19-1889082",
                "19-1893640",
                "18-1877852",
                "19-1888244",
                "18-1878365",
                "19-1900681",
                "18-1879410",
                "18-1875318",
                "19-1887083",
                "18-1884403",
                "19-1889243",
                "19-1887338",
                "19-1891032",
                "19-1901851",
                "18-1883044",
                "19-1889233",
                "19-1903682",
                "19-1899811",
                "19-1897010",
                "18-1877393",
                "18-1885271",
                "19-1897316",
                "19-1890288",
                "18-1881047",
                "19-1889437",
                "19-1889602",
                "18-1885628",
                "19-1899391",
                "19-1894334",
                "19-1901469",
                "18-1881300",
                "19-1902354",
                "18-1876909",
                "19-1896586",
                "19-1887720",
                "19-1897786",
                "19-1902907",
                "19-1901079",
                "18-1883903",
                "19-1888182",
                "19-1900105",
                "19-1894423",
                "19-1896703",
                "18-1883549",
                "18-1882346",
                "18-1885689",
                "18-1875954",
                "19-1893156",
                "19-1893512",
                "19-1889145",
                "19-1901736",
                "19-1894129",
                "19-1895232",
                "18-1886647",
                "18-1883477",
                "18-1879072",
                "19-1900153",
                "19-1893669",
                "19-1899898",
                "19-1901463",
                "19-1894322",
                "19-1893880",
                "19-1897891",
                "19-1886837",
                "18-1880425",
                "19-1898944",
                "18-1883714",
                "18-1884473",
                "19-1901311",
                "18-1876495",
                "19-1889689",
                "19-1901349",
                "18-1879630",
                "19-1899897",
                "19-1887318",
                "18-1883602",
                "18-1873536",
                "18-1885927",
                "18-1875904",
                "19-1900355",
                "18-1876088",
                "19-1896138",
                "19-1890847",
                "19-1890917",
                "18-1873834",
                "19-1888274",
                "18-1885695",
                "19-1893157",
                "19-1896343",
                "18-1883306",
                "19-1888655",
                "19-1894215",
                "19-1897351",
                "19-1896394",
                "19-1901010",
                "18-1877645",
                "19-1891483",
                "19-1898078",
                "18-1883641",
                "19-1893297",
                "19-1889857",
                "18-1875093",
                "19-1899355",
                "19-1902169",
                "19-1899198",
                "18-1875181",
                "18-1876392",
                "19-1886864",
                "18-1878062",
                "19-1900615",
                "19-1899434",
                "19-1898777",
                "19-1901886",
                "18-1881109",
                "19-1899167",
                "18-1884148",
                "18-1884606",
                "18-1874910",
                "19-1902837",
                "18-1878409",
                "19-1897538",
                "18-1880649",
                "17-1854674",
                "18-1882885",
                "19-1894100",
                "19-1888904",
                "18-1886591",
                "19-1888662",
                "19-1895684",
                "18-1885081",
                "18-1872654",
                "19-1894054",
                "19-1898310",
                "19-1903196",
                "18-1881915",
                "19-1896189",
                "18-1885121",
                "19-1886803",
                "19-1897879",
                "18-1886124",
                "18-1879478",
                "18-1884222",
                "18-1878157",
                "18-1884339",
                "19-1896473",
                "18-1879745",
                "18-1874351",
                "18-1885128",
                "19-1899693",
                "19-1887124",
                "19-1896415",
                "19-1896989",
                "18-1879553",
                "19-1898382",
                "19-1900443",
                "19-1902141",
                "19-1892286",
                "18-1883455",
                "18-1880988",
                "18-1882543",
                "19-1888053",
                "18-1878388",
                "19-1889594",
                "18-1877632",
                "18-1881256",
                "18-1878881",
                "19-1889047",
                "19-1894418",
                "18-1885997",
                "19-1897621",
                "19-1899361",
                "18-1883328",
                "19-1891821",
                "19-1898581",
                "19-1899377",
                "19-1899018",
                "19-1889674",
                "19-1897128",
                "19-1889396",
                "18-1878939",
                "19-1894831",
                "18-1885214",
                "18-1876779",
                "19-1891125",
                "19-1892582",
                "19-1896246",
                "18-1877796",
                "19-1900134",
                "19-1898362",
                "19-1890202",
                "18-1879697",
                "19-1898789",
                "19-1894213",
                "19-1893765",
                "19-1893884",
                "19-1901620",
                "19-1898579",
                "18-1871850",
                "19-1892263",
                "18-1875552",
                "18-1877810",
                "18-1882425",
                "18-1884718",
                "19-1888902",
                "18-1880803",
                "18-1884478",
                "19-1888180",
                "18-1883629",
                "18-1875932",
                "19-1890593",
                "18-1882874",
                "19-1888869",
                "19-1899447",
                "19-1901661",
                "18-1880882",
                "19-1897044",
                "18-1876768",
                "19-1899413",
                "19-1890182",
                "18-1877853",
                "18-1883267",
                "18-1882785",
                "19-1889352",
                "18-1875556",
                "19-1895951",
                "19-1888446",
                "19-1900492",
                "19-1897068",
                "18-1874726",
                "19-1892097",
                "18-1877226",
                "19-1901000",
                "18-1879033",
                "19-1887170",
                "19-1899497",
                "19-1893221",
                "18-1886740",
                "19-1902408",
                "19-1890985",
                "19-1893158",
                "19-1898994",
                "19-1898557",
                "19-1896073",
                "18-1883874",
                "19-1893017",
                "18-1875446",
                "18-1884768",
                "19-1897754",
                "18-1884531",
                "19-1900126",
                "19-1891299",
                "18-1874286",
                "19-1896590",
                "18-1884375",
                "19-1895104",
                "19-1893182",
                "19-1889426",
                "19-1887662",
                "19-1891078",
                "19-1891442",
                "19-1887168",
                "18-1870552",
                "19-1886970",
                "19-1889308",
                "19-1901383",
                "18-1875585",
                "18-1882035",
                "19-1892374",
                "18-1873511",
                "19-1893241",
                "19-1898879",
                "18-1877498",
                "19-1890062",
                "18-1880212",
                "19-1889311",
                "18-1885114",
                "18-1884248",
                "19-1889079",
                "19-1899174",
                "18-1878912",
                "19-1888235",
                "18-1882423",
                "19-1894396",
                "18-1875886",
                "18-1879784",
                "19-1895614",
                "18-1881852",
                "18-1883298",
                "19-1889133",
                "19-1889064",
                "19-1887900",
                "19-1898527",
                "19-1890400",
                "19-1889761",
                "19-1896568",
                "18-1875915",
                "18-1876075",
                "18-1882930",
                "18-1885861",
                "18-1881761",
                "19-1893939",
                "18-1880436",
                "19-1889336",
                "19-1899729",
                "19-1899369",
                "18-1880763",
                "19-1896601",
                "19-1893174",
                "19-1892909",
                "18-1877815",
                "18-1878249",
                "19-1890093",
                "19-1891440",
                "19-1894186",
                "19-1898299",
                "19-1893005",
                "19-1889347",
                "19-1895134",
                "19-1893972",
                "19-1898742",
                "19-1902212",
                "18-1883013",
                "19-1900391",
                "18-1884054",
                "18-1886598",
                "18-1877026",
                "18-1881083",
                "18-1886031",
                "18-1872020",
                "18-1875765",
                "18-1883147",
                "19-1888766",
                "18-1881722",
                "19-1889457",
                "19-1898433",
                "19-1893475",
                "19-1889317",
                "19-1892529",
                "18-1870362",
                "19-1899571",
                "18-1873250",
                "18-1879093",
                "19-1888730",
                "19-1898432",
                "19-1899568",
                "19-1892096",
                "18-1876826",
                "19-1900877",
                "18-1883977",
                "19-1889940",
                "18-1879505",
                "18-1883939",
                "18-1881005",
                "19-1897728",
                "19-1892894",
                "18-1886465",
                "19-1896935",
                "18-1877504",
                "19-1893004",
                "19-1890199",
                "19-1893954",
                "19-1891898",
                "19-1895495",
                "19-1901202",
                "19-1900044",
                "18-1883462",
                "18-1877368",
                "19-1889458",
                "19-1901136",
                "18-1862806",
                "18-1882745",
                "19-1887109",
                "18-1882495",
                "19-1889982",
                "18-1886138",
                "19-1895094",
                "18-1873725",
                "19-1889224",
                "19-1887048",
                "19-1894393",
                "18-1877869",
                "19-1888893",
                "18-1875202",
                "19-1903265",
                "18-1878323",
                "19-1900438",
                "18-1878367",
                "19-1895941",
                "19-1894546",
                "19-1898397",
                "19-1895965",
                "18-1884441",
                "19-1902179",
                "19-1901352",
                "19-1886967",
                "18-1879687",
                "18-1886654",
                "19-1892078",
                "19-1888470",
                "18-1885154",
                "18-1883735",
                "18-1876995",
                "18-1881457",
                "19-1901158",
                "19-1896236",
                "19-1900622",
                "18-1871744",
                "19-1888617",
                "18-1876844",
                "19-1888366",
                "19-1889855",
                "18-1882845",
                "19-1895480",
                "19-1889586",
                "19-1896991",
                "18-1883726",
                "19-1900029",
                "18-1880468",
                "19-1889226",
                "18-1883045",
                "18-1884395",
                "19-1897668",
                "19-1892125",
                "18-1882554",
                "19-1903608",
                "18-1886361",
                "19-1890615",
                "18-1878385",
                "18-1884105",
                "19-1887363",
                "18-1875598",
                "19-1887952",
                "19-1897386",
                "18-1885971",
                "18-1881116",
                "19-1896088",
                "19-1890252",
                "18-1878285",
                "19-1888359",
                "19-1893682",
                "19-1892725",
                "19-1889551",
                "18-1879380",
                "19-1897818",
                "19-1899603",
                "19-1894384",
                "19-1900689",
                "19-1890901",
                "19-1891189",
                "18-1878699",
                "18-1881955",
                "19-1886796",
                "18-1880936",
                "19-1886863",
                "18-1880373",
                "18-1877299",
                "19-1898401",
                "19-1891760",
                "18-1882923",
                "19-1895498",
                "18-1886028",
                "18-1884059",
                "18-1881055",
                "19-1900469",
                "19-1896315",
                "19-1896530",
                "18-1882661",
                "18-1866888",
                "18-1878543",
                "18-1882802",
                "18-1883120",
                "18-1879374",
                "19-1897903",
                "19-1899463",
                "18-1882603"
                ]


@app.route("/IOIempQuarterly", methods=['GET', 'POST'])
def upload_IOIempQuarterly():
    if request.method == 'POST':
        print(request.files['file'])
        f = request.files['file']
        
        test = pd.read_excel(f)
        
        test.fillna('',inplace=True)
        
        
        #Cleaning
        if test.iloc[0][0] == '':
            data_xls = pd.read_excel(f,skiprows=2)
        else:
            data_xls = pd.read_excel(f)
        
        last7 = data_xls['Matter/Case ID#'].apply(lambda x: x[3:])
        CaseNum = data_xls['Matter/Case ID#']
        data_xls['Temp Hyperlinked Case #']='=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseNum +'"' +')'
        move=data_xls['Temp Hyperlinked Case #']
        data_xls.insert(0,'Hyperlinked Case #', move)           
        del data_xls['Temp Hyperlinked Case #']
        
        data_xls.fillna('',inplace=True)
        
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Bronx Legal Services','BxLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Brooklyn Legal Services','BkLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Queens Legal Services','QLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Manhattan Legal Services','MLS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Staten Island Legal Services','SILS')
        data_xls['Assigned Branch/CC'] = data_xls['Assigned Branch/CC'].str.replace('Legal Support Unit','LSU')
        
        
        """
        
        #Determining 'level of service' from 3 fields       
        def HRA_Service_Type(Employment_Tier):
            Employment_Tier = str(Employment_Tier)
            
            if Employment_Tier.startswith("Advice-No Retainer") == True:
                return 'B'
            elif Employment_Tier.startswith("UI Representation") == True:
                return 'T1'
            elif Employment_Tier.startswith("Advice-Investigation Retainer") == True:
                return 'T1'
            elif Employment_Tier.startswith("Demand Letter-Negotiation") == True:
                return 'T1'
            elif Employment_Tier.startswith("Admin Rep") == True:
                return 'T2'
            elif Employment_Tier.startswith("Litigation") == True:
                return 'T2'
            else:
                return '***Needs Cleanup***'
        data_xls['HRA Service Type'] = data_xls.apply(lambda x: HRA_Service_Type(x['HRA IOI Employment Law IOI Employment Tier Category:']), axis=1)

        data_xls['HRA Proceeding Type'] = 'EMP'
        
        """
        
        #Putting Employment Work in HRA Baskets
        
        def HRA_Case_Coding(LoS,LPC,SLPC,Retainer):
            if LoS.startswith('Advice') and Retainer.startswith('Investigation'):
                return "T1-EMPOTH_"
            elif LoS.startswith('Brief') and Retainer.startswith('Investigation'):
                return "T1-EMPOTH_"
            elif LoS.startswith('Advice') and SLPC.startswith('295'):
                return "B -RAP"
            elif LoS.startswith('Brief') and SLPC.startswith('295'):
                return "B -RAP"
            elif LoS.startswith('Advice') and SLPC.startswith('296'):
                return "B -CRD"
            elif LoS.startswith('Brief') and SLPC.startswith('296'):
                return "B -CRD"   
            elif LoS.startswith('Advice') and SLPC.startswith('297'):
                return "B -CGC"
            elif LoS.startswith('Brief') and SLPC.startswith('297'):
                return "B -CGC"
            elif LoS.startswith('Advice') or LoS.startswith('Brief'):
                return "B -EMP"
            elif LoS.startswith('Out-of-Court'):
                return "T1-PRELIT"
            elif LoS.startswith('Rep') and LPC.startswith('76'):
                return "T1-UIWC"
            elif LoS.startswith('Rep') and LPC.startswith('93'):
                return "T1-LICE"
            elif LoS.startswith('Rep') and SLPC.startswith('22-1'):
                return "T2-WAGE"
            elif LoS.startswith('Rep') and SLPC.startswith('22-2'):
                return "T2-WAGE"    
            elif LoS.startswith('Rep') and SLPC.startswith('22-3'):
                return "T2-WAGE"
            elif LoS.startswith('Rep') and LPC.startswith('01 Bank'):
                return "T2-BANK"
            elif LoS.startswith('Representation') and LPC == '21 Employment Discrimination':
                return "T2-DIS"
            elif LoS.startswith('Rep') and SLPC.startswith('293'):
                return "T2-FML"
            elif LoS.startswith('Rep') and SLPC.startswith('22-4'):
                return "T2-WB"
            elif LoS.startswith('Rep') and SLPC.startswith('22-5'):
                return "T2-WB"
            elif LoS.startswith('Rep'):
                return "T1-EMPOTH_"
            elif SLPC == "Mandamus Action": 
                return 'T2-FED'
            elif SLPC == "EOIR-26" or SLPC == "EOIR-27":
                return 'T2-FED'
            else:
                return "***Needs Cleanup***"
        
        data_xls['HRA_Case_Coding'] = data_xls.apply(lambda x: HRA_Case_Coding(x['Level of Service'],x['Legal Problem Code'],x['Special Legal Problem Code'],x['HRA IOI Employment Law Retainer?']), axis = 1)
        

        #Does case need special legal problem code?
        def SPLC_problem(LoS,SPLC,HRACode):
            if LoS == 'Hold For Review':
                return SPLC
            elif HRACode.startswith('***') and LoS != "":
                return '***Needs SPLC***'
            else:
                return SPLC
                
        data_xls['Special Legal Problem Code'] = data_xls.apply(lambda x: SPLC_problem(x['Level of Service'],x['Special Legal Problem Code'],x['HRA_Case_Coding']), axis = 1)
        
        
        #Can case be reported based on income?
        
        def Income_Exclude(IncomePct,Waiver):
            IncomePct = int(IncomePct)
            Waiver = str(Waiver)
            if IncomePct > 200 and Waiver.startswith('Y') == False:
                return 'Needs Income Waiver'
            else:
                return ''
        

        data_xls['Exclude due to Income?'] = data_xls.apply(lambda x: Income_Exclude(x['Percentage of Poverty'],x['HRA IOI Employment Law If Client is Over 200% of FPL, Did you seek a waiver from HRA?']), axis=1)
        
        data_xls['Open Month'] = data_xls['Date Opened'].apply(lambda x: str(x)[:2])
        data_xls['Open Day'] = data_xls['Date Opened'].apply(lambda x: str(x)[3:5])
        data_xls['Open Year'] = data_xls['Date Opened'].apply(lambda x: str(x)[6:])
        data_xls['Open Construct'] = data_xls['Open Year'] + data_xls['Open Month'] + data_xls['Open Day']
        
        #DHCI form needed?
        
        def DHCI_Needed(DHCI,LoS,Open_Construct):
            if LoS.startswith('Advice'):
                return ''
            elif LoS.startswith('Brief'):
                return ''
            elif Open_Construct == 'na':
                return ''
            elif int(Open_Construct) < 20181115:
                return ''
            elif DHCI != 'Yes':
                return 'Needs DHCI'
            else:
                return ''
        
        data_xls['Needs DHCI?'] = data_xls.apply(lambda x: DHCI_Needed(x['HRA IOI Employment Law DHCI Form?'],x['Level of Service'],x['Open Construct']), axis=1)
        
         
        #Manipulable Dates               
        
        def Eligibility_Date(Effective_Date,Date_Opened):
            if Effective_Date != '':
                return Effective_Date
            else:
                return Date_Opened
        data_xls['Eligibility_Date'] = data_xls.apply(lambda x : Eligibility_Date(x['IOI HRA Effective Date (optional) (IOI 2)'],x['Date Opened']), axis = 1)
        
        data_xls['Open Month'] = data_xls['Eligibility_Date'].apply(lambda x: str(x)[:2])
        data_xls['Open Day'] = data_xls['Eligibility_Date'].apply(lambda x: str(x)[3:5])
        data_xls['Open Year'] = data_xls['Eligibility_Date'].apply(lambda x: str(x)[6:])
        data_xls['Open Construct'] = data_xls['Open Year'] + data_xls['Open Month'] + data_xls['Open Day']
        
        data_xls['Subs Month'] = data_xls['HRA IOI Employment Law HRA Date Substantial Activity Performed'].apply(lambda x: str(x)[:2])
        data_xls['Subs Day'] = data_xls['HRA IOI Employment Law HRA Date Substantial Activity Performed'].apply(lambda x: str(x)[3:5])
        data_xls['Subs Year'] = data_xls['HRA IOI Employment Law HRA Date Substantial Activity Performed'].apply(lambda x: str(x)[6:])
        data_xls['Subs Construct'] = data_xls['Subs Year'] + data_xls['Subs Month'] + data_xls['Subs Day']
        data_xls['Subs Construct'] = data_xls.apply(lambda x : x['Subs Construct'] if x['Subs Construct'] != '' else 0, axis = 1)
        
        #Substantial Activity for Rollover?
                
        
        def Needs_Rollover(Open_Construct,Substantial_Activity,Substantial_Activity_Date,CaseID,ReportedFY19):
            if int(Open_Construct) >= 20190701:
                return ''
            elif Substantial_Activity != '' and int(Substantial_Activity_Date) >= 20190701 and int(Substantial_Activity_Date) <= 20200630:
                return ''
            elif CaseID in ReportedFY19:
                return 'Needs Substantial Activity in FY20'
            else:
                return ''
                
        data_xls['Needs Substantial Activity?'] = data_xls.apply(lambda x: Needs_Rollover(x['Open Construct'],x['HRA IOI Employment Law HRA Substantial Activity'],x['Subs Construct'],x['Matter/Case ID#'], ReportedFY19), axis=1)
        
        
        #Assign Outcomes
        def AdviceOutcomeDate(HRAOutcome,HRAOutcomeDate,DateClosed):
            if HRAOutcomeDate == '' and HRAOutcome == 'Advice Given':
                return DateClosed
            else:
                return HRAOutcomeDate
                
        data_xls['HRA IOI Employment Law HRA Outcome Date:'] = data_xls.apply(lambda x: AdviceOutcomeDate(x['HRA IOI Employment Law HRA Outcome:'],x['HRA IOI Employment Law HRA Outcome Date:'],x['Date Closed']), axis = 1)
        
        def AdviceOutcome(HRAOutcome,Employment_Tier,CaseDisposition,HRAOutcomeDate):
            if HRAOutcome == '' and CaseDisposition== 'Closed' and Employment_Tier == 'Advice-No Retainer':
                return 'Advice Given'
            elif HRAOutcome == '' and CaseDisposition == 'Closed':
                return '**Needs Outcome**'
            elif HRAOutcome != '' and CaseDisposition == 'Closed'and HRAOutcomeDate == '':
                return '**Needs Outcome Date**'
            else:
                return HRAOutcome

        data_xls['HRA Outcome'] = data_xls.apply(lambda x: AdviceOutcome(x['HRA IOI Employment Law HRA Outcome:'], x['HRA IOI Employment Law IOI Employment Tier Category:'], x['Case Disposition'],x['HRA IOI Employment Law HRA Outcome Date:']), axis = 1)


        #Better names & HRA Names

        data_xls['Employment Tier Category'] = data_xls['HRA IOI Employment Law IOI Employment Tier Category:']
        
        data_xls['Client Name'] = data_xls['Full Person/Group Name (Last First)']
        
        data_xls['Office'] = data_xls['Assigned Branch/CC']
        
        data_xls['Unique_ID'] = 'LSNYC'+data_xls['Matter/Case ID#']
        
        data_xls['Last_Initial'] = data_xls['Client Last Name'].str[1]
        data_xls['First_Initial'] = data_xls['Client First Name'].str[1]
        
        data_xls['Year_of_Birth'] = data_xls['Date of Birth'].str[-4:]
        
        #gender
        def HRAGender (gender):
            if gender == 'Male' or gender == 'Female':
                return gender
            else:
                return 'Other'
        data_xls['Gender'] = data_xls.apply(lambda x: HRAGender(x['Gender']), axis=1)
        
        data_xls['Country of Origin'] = ''
        
        #county=borough
        data_xls['Borough'] = data_xls['County of Residence']
        
        #household size etc.
        data_xls['Household_Size'] = data_xls['Number of People under 18'].astype(int) + data_xls['Number of People 18 and Over'].astype(int)
        data_xls['Number_of_Children'] = data_xls['Number of People under 18']
        
        #Income Eligible?
        data_xls['Annual_Income'] = data_xls['Total Annual Income ']
        def HRAIncElig (PercentOfPoverty):
            if PercentOfPoverty > 200:
                return 'NO'
            else:
                return 'YES'
        data_xls['Income_Eligible'] = data_xls.apply(lambda x: HRAIncElig(x['Percentage of Poverty']), axis=1)
        
        def IncWaiver (eligible,waiverdate):
            if eligible == 'NO' and waiverdate != '':
                return 'Income'
            else:
                return ''
        data_xls['Waiver_Type'] = data_xls.apply(lambda x: IncWaiver(x['Income_Eligible'],x['HRA IOI Employment Law If Client is Over 200% of FPL, Did you seek a waiver from HRA?']), axis=1)
        
        def IncWaiverDate (waivertype,date):
            if waivertype != '':
                return date
            else:  
                return ''
                
        data_xls['Waiver_Approval_Date'] = data_xls.apply(lambda x: IncWaiverDate(x['Waiver_Type'],x['HRA IOI Employment Law Income Waiver Date']), axis = 1)
        

        #Other Cleanup
        
        
        data_xls['Referral_Source'] = 'None'
        
        def ProBonoCase (branch, pai):
            if branch == "LSU" or pai == "Yes":
                return "YES"
            else:
                return "NO"
                
        def PriorEnrollment (casenumber):
            if casenumber in ReportedFY19:
                return 'FY 19'
                
        data_xls['Prior_Enrollment_FY'] = data_xls.apply(lambda x:PriorEnrollment(x['Matter/Case ID#']), axis = 1)
                
        data_xls['Pro_Bono'] = data_xls.apply(lambda x:ProBonoCase(x['Assigned Branch/CC'], x['PAI Case?']), axis = 1)
       
        data_xls['Service_Type_Code'] = data_xls['HRA_Case_Coding'].apply(lambda x: x[:2] if x != 'Hold For Review' else '')

        data_xls['Proceeding_Type_Code'] = data_xls['HRA_Case_Coding'].apply(lambda x: x[3:] if x != 'Hold For Review' else '')
        
        data_xls['Outcome'] = data_xls['HRA IOI Employment Law HRA Outcome:']
        data_xls['Outcome_Date'] = data_xls['HRA IOI Employment Law HRA Outcome Date:']
        
        data_xls['Seized_at_Border'] = ''
        
        data_xls['Group'] = ''
        
        
        
        #sorting by borough and advocate
        data_xls = data_xls.sort_values(by=['Office','Primary Advocate'])
        
        
        
        #REPORTING VERSION Put everything in the right order
        data_xls = data_xls[['Unique_ID','Last_Initial','First_Initial','Year_of_Birth','Gender','Country of Origin','Borough','Zip Code','Language','Household_Size','Number_of_Children','Annual_Income','Income_Eligible','Waiver_Type','Waiver_Approval_Date','Eligibility_Date','Referral_Source','Service_Type_Code','Proceeding_Type_Code','Outcome','Outcome_Date','Seized_at_Border','Group','Prior_Enrollment_FY','Pro_Bono','Hyperlinked Case #','Office','Primary Advocate','Client Name','Level of Service','Legal Problem Code','Special Legal Problem Code','HRA_Case_Coding','Exclude due to Income?','Needs DHCI?','Needs Substantial Activity?']]
        
        
        
        
        output_filename = f.filename     
        writer = pd.ExcelWriter("app\\sheets\\"+output_filename, engine = 'xlsxwriter')
        data_xls.to_excel(writer, sheet_name='Sheet1',index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        link_format = workbook.add_format({'font_color':'blue', 'bold':True, 'underline':True})
        problem_format = workbook.add_format({'bg_color':'yellow'})
        
        
        
        worksheet.set_column('C:BL',30)
        worksheet.conditional_format('E1:E100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
        worksheet.conditional_format('F1:F100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"***Needs SPLC***"',
                                                 'format': problem_format})
        worksheet.conditional_format('G1:G100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Income Waiver"',
                                                 'format': problem_format})
        worksheet.conditional_format('H1:H100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs DHCI"',
                                                 'format': problem_format})
        worksheet.conditional_format('I1:I100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"Needs Substantial Activity in FY20"',
                                                 'format': problem_format})
        worksheet.conditional_format('J1:J100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '""',
                                                 'format': problem_format})
        worksheet.conditional_format('K1:K100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"**Needs Outcome**"',
                                                 'format': problem_format})
        worksheet.conditional_format('K1:K100000',{'type': 'cell',
                                                 'criteria': '==',
                                                 'value': '"**Needs Outcome Date**"',
                                                 'format': problem_format})
        writer.save()
        
        return send_from_directory('sheets',output_filename, as_attachment = True, attachment_filename = "Formatted " + f.filename)

    return '''
    <!doctype html>
    <title>IOI Employment Quarterly</title>
    <link rel="stylesheet" href="/static/css/main.css">
    <h1>Quarterly Formatter for IOI Employment Cases:</h1>
    <form action="" method=post enctype=multipart/form-data>
    <p><input type=file name=file><input type=submit value=IOI-ify!>
    </form>
    <h3>Instructions:</h3>
    <ul type="disc">
    <li>This tool is meant to be used in conjunction with the LegalServer report called <a href="https://lsnyc.legalserver.org/report/dynamic?load=2020" target="_blank">"Grants Management IOI Employment (3474) Report"</a>.</li>
    <li>Browse your computer using the field above to find the LegalServer excel document that you want to process for IOI.</li> 
    <li>Once you have identified this file, click ‘IOI-ify!’ and you should shortly be given a prompt to either open the file directly or save the file to your computer.</li> 
    <li>When you first open the file, all case numbers will display as ‘0’ until you click “Enable Editing” in excel, this will populate the fields.</li> </ul>
    </br>
    <a href="/">Home</a>
    '''
