#General Purpose Functions to be used in LSNYC Report Prep 
import datetime

#Remove extraneous summary rows
def RemoveNoCaseID(CaseID):
        if CaseID == '' or CaseID == 'nan' or CaseID.startswith("Unique") == True:
            return 'No Case ID'

        else:
            return str(CaseID)


#Turn a case ID# into a hyperlink
def Hyperlinker(CaseID):
        last7 = CaseID[-7:]
        hyperlinkedID = '=HYPERLINK("https://lsnyc.legalserver.org/matter/dynamic-profile/view/'+last7+'",'+ '"' + CaseID +'"' +')'
        return hyperlinkedID 
        

#Take a date formatted as MM/DD/YYYY and make it YYYYMMDD so it can be easily compared to other dates           
def DateMaker (Date):
            
            if isinstance(Date,datetime.datetime) == True:
                Date = Date.strftime("%m/%d/%Y")
            DateMonth = Date[:2]
            DateDay = Date[3:5]
            DateYear = Date[6:]
            
            
            if Date == "":
                return ""            
            else:
                return int(DateYear + DateMonth + DateDay)


#Turn the Queens micro-cities into just saying 'queens'       
QueensNeighborhoods = ["Arverne",
                        "Astoria",
                        "Bayside",
                        "Bellerose",
                        "Cambria Heights",
                        "College Point",
                        "Corona",
                        "Elmhurst",
                        "Far Rockaway",
                        "Flushing",
                        "Forest Hills",
                        "Fresh Meadows",
                        "Glendale",
                        "Hollis",
                        "Jackson Heights",
                        "Jackson Hts",
                        "Jamaica",
                        "Kew Gardens",
                        "Laurelton",
                        "Little Neck",
                        "Long Is City",
                        "Long Island City",
                        "Maspeth",
                        "Middle Village",
                        "Ozone Park",
                        "Queens",
                        "Queens Village",
                        "Rego Park",
                        "Richmond Hill",
                        "Ridgewood",
                        "Rockaway Beach",
                        "Rockaway Park",
                        "Rosedale",
                        "S Ozone Park",
                        "Saint Albans",
                        "South Ozone Park",
                        "South Richmond Hill",
                        "Springfield Gardens",
                        "Sunnyside",
                        "Whitestone",
                        "Woodhaven",
                        "Woodside"]
       
def QueensConsolidater(City):
    if City in QueensNeighborhoods:
        return "Queens"
    else:
        return City
        
#Borough office Abbreviators

def OfficeAbbreviator(AssignedBranch):
    if AssignedBranch == "Bronx Legal Services":
        return "BxLS"
    elif AssignedBranch == "Brooklyn Legal Services":
        return "BkLS"
    elif AssignedBranch == "Queens Legal Services":
        return "QLS"    
    elif AssignedBranch == "Manhattan Legal Services":
        return "MLS"    
    elif AssignedBranch == "Staten Island Legal Services":
        return "SILS"    
    elif AssignedBranch == "Legal Support Unit":
        return "LSU"    
    else:
        return AssignedBranch    
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        