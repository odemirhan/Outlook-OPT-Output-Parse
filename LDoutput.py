import win32com.client
import pyodbc
import pandas as pd
from datetime import datetime


conn=pyodbc.connect('Driver={SQL Server};'
                      'Server=10.1.0.36;'
                      'Database=FuelMasterDB;'
                      'Integrated Security=False;'
                      'uid=sa;'
                        'pwd=Fwd7ygh52.;'
                        
                              )
outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
          
TOfolder= outlook.GetDefaultFolder(6).Folders.Item("Landing Output")

ArchieveLD= outlook.GetDefaultFolder(6).Folders.Item("LD Archieve")
messages = TOfolder.Items

def isdate(datee, timee):
    try:
        timee=timee.replace("AIRPORT","")
        timee=timee.replace("POLICY","")
        dt=datetime.strptime(datee+" "+timee, "%d%b%y %H:%M:%S")
        return True
    except:
        return False


print(len(messages))
for cntm in range(len(messages)):
    try:
        if messages[cntm+1].Class==43:
           if messages[cntm+1].SenderEmailType=='EX':
               sender=messages[cntm+1].Sender.GetExchangeUser().PrimarySmtpAddress
           else:
               sender=messages[cntm+1].SenderEmailAddress
    except:
        sender=""
    try:
        Daterec=messages[cntm+1].ReceivedTime
    except:
        Daterec=""
    
    
    if len(sender)>50:
        sender=sender[0:45]
    try:    
        mailsub=messages[cntm+1].Subject
    except:
        mailsub=""
    
    if 'SIM' in mailsub or 'Sim' in mailsub or 'sim' in mailsub:
        print('sildim')
    else:
        smailsub=mailsub.split(" ")
        aircrafts=[x for x in smailsub if x.startswith('TC') or x.startswith('9H')]
        
        try:
            aircraft=aircrafts[0]
        except:
            aircraft=""

        try:
            wholebody=messages[cntm+1].Body
            if 'SIM' in wholebody or 'Sim'  in wholebody or 'sim' in wholebody:
                print('sildim2')
        except:
            wholebody=""

        else:
            wholebody=wholebody.replace('\n',' ')
            splittedbody=wholebody.split(" ")
            splittedbody=[x for x in splittedbody if x != '']
            splittedbody =[x.replace('\n','') for x in splittedbody]
            splittedbody =[x.replace('\r','') for x in splittedbody]
            
            CALCDATE=""
            AIRPORT=""
            RUNWAY=""
            LDA=""
            CONDITION=""
            WIND=""
            OAT=""
            QNH=""
            PACKS=""
            LDCLIMGRAD=""
            APRCLIMGRAD=""
            JAROPSCLIMGRAD=""
            ENLDWEIGHT=""
            ENFLAPS=""
            VREF=""
            MM=""
            AB1=""
            AB2=""
            AB3=""
            MA=""
            NNC=""
            AIRPORTELEV=""
              
            AIRPORTDB=""
            POLICYDB=""
            MELCDL=""


            for items in range(len(splittedbody)):

                try:
                    if "PERFORMANCE" in splittedbody[items]:
                        timee=splittedbody[items+2]
                        timee=timee.replace("AIRPORT","")
                        dt=splittedbody[items+1]+" "+timee
                        if isdate(splittedbody[items+1], timee):
                            CALCDATE=datetime.strptime(dt, "%d%b%y %H:%M:%S")
                except:
                    CALCDATE=""
                try:
                    if "AIRPORT" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            AIRPORT=splittedbody[items+2]
                except:
                    AIRPORT=""
                try:
                        
                    if "RUNWAY" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            RUNWAY=splittedbody[items+2]
                except:
                    RUNWAY=""
                try:
                        
                    if "LDA:" in splittedbody[items]:
                        LDA=splittedbody[items+1]
                except:
                    LDA=""
                try:
                    
                    if "CONDITION" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            CONDITION=splittedbody[items+2]
                except:
                    CONDITION=""
                try:
                        
                    if "WIND" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            WIND=splittedbody[items+2]
                except:
                    WIND=""
                try:
                    if "OAT" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            OAT=splittedbody[items+2]
                except:
                    OAT=""
                try:
                    if "QNH" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            QNH=splittedbody[items+2]
                except:
                    QNH=""
                try:
                    if "PACKS" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            PACKS=splittedbody[items+2]
                except:
                    PACKS=""


                try:
                        
                    if "LANDING" in splittedbody[items]:
                        if splittedbody[items+1]=="CLIMB":
                            LDCLIMGRAD=splittedbody[items+3]
                except:
                    LDCLIMGRAD=""
                try:
                        
                    if "APPROACH" in splittedbody[items]:
                        if splittedbody[items+1]=="CLIMB":
                            APRCLIMGRAD=splittedbody[items+3]
                except:
                    APRCLIMGRAD=""
                try:
                        
                    if "REQ" in splittedbody[items]:
                        if splittedbody[items+1]=="JAROPS":
                            JAROPSCLIMGRAD=splittedbody[items+3]
                except:
                    JAROPSCLIMGRAD=""

                try:
                    if "WEIGHT" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            ENLDWEIGHT=splittedbody[items+3]
             
                except:
                    ENLDWEIGHT=""
                try:
                    if "FLAPS" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            ENFLAPS=splittedbody[items+3]
             
                except:
                    ENFLAPS=""
                try:
                    if "VREF" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            VREF=splittedbody[items+3]
             
                except:
                    VREF=""

            
                if "REQD" in splittedbody[items]:
                    inter=items
                    
                    
                    for cntcnt in range(len(splittedbody)-inter-3):
                        try:
                            if splittedbody[items+cntcnt]=="MANUAL)":
                                MM=splittedbody[items+cntcnt-3]
                        except:
                            MM=""
                        try:
                            if splittedbody[items+cntcnt]=="1)":
                                AB1=splittedbody[items+cntcnt-4]
                        except:
                            AB1=""
                        try:
                            if splittedbody[items+cntcnt]=="2)":
                                AB2=splittedbody[items+cntcnt-4]
                        except:
                            AB2=""
                        try:
                            if splittedbody[items+cntcnt]=="3)":
                                AB3=splittedbody[items+cntcnt-4]
                        except:
                            AB3=""
                        try:
                            if "AUTO)" in splittedbody[items+cntcnt]:
                                MA=splittedbody[items+cntcnt-3]
                        except:
                            MA=""
                try:
                    if "NNC" in splittedbody[items]:
                        NNC=splittedbody[items+1]
             
                except:
                    NNC=""
               
                
                try:
                    
                    if "AIRPORT" in splittedbody[items]:
                        if "ELEV" in splittedbody[items+1]:
                            AIRPORTELEV=splittedbody[items+2]
                except:
                    AIRPORTELEV=""
                
               
                try:
                        
                    if "AIRPORTDB" in splittedbody[items]:
                        AIRPORTDB=splittedbody[items+1]
                except:
                    AIRPORTDB=""
                try:
                        
                    if "POLICYDB" in splittedbody[items]:
                        POLICYDB=splittedbody[items+1]
                except:
                    POLICYDB=""
                try:
                    
                    if "MEL/CDL" in splittedbody[items]:
                        MELCDL=splittedbody[items+1]
                except:
                    MELCDL=""

            listtowrite=(sender, Daterec, CALCDATE, aircraft,
            AIRPORT,
            RUNWAY,
            LDA,
            CONDITION,
            WIND,
            OAT,
            QNH,
            PACKS,
            LDCLIMGRAD,
            APRCLIMGRAD,
            JAROPSCLIMGRAD,
            ENLDWEIGHT,
            ENFLAPS,
            VREF,
            MM,
            AB1,
            AB2,
            AB3,
            MA,
            NNC,
            AIRPORTELEV,
            AIRPORTDB,
            POLICYDB,
            MELCDL)
            
            print(listtowrite)
            try:
                cur=conn.cursor()
                cur.execute("INSERT INTO dbo.[OPT Landing] VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                listtowrite)
                conn.commit()
            except Exception as E:
                print(E)
                pass
    try:
        messages[cntm+1].Move(ArchieveLD)
    except:
        pass
            
