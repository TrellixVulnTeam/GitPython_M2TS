#General Purpose Functions to be used in LSNYC Report Prep 

#Remove extraneous summary rows
def RemoveNoCaseID(CaseID):
        if CaseID == '' or CaseID == 'nan':
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
            DateMonth = Date[:2]
            DateDay = Date[3:5]
            DateYear = Date[6:]
            if Date == "":
                return ""            
            else:
                return int(DateYear + DateMonth + DateDay)