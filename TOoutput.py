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
          
TOfolder= outlook.GetDefaultFolder(6).Folders.Item("Takeoff Output")

ArchieveTO= outlook.GetDefaultFolder(6).Folders.Item("TO Archieve")
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
            CONDITION=""
            WIND=""
            OAT=""
            QNH=""
            PACKS=""
            TOW=""
            MFHR=""
            MLOH=""
            FLAPS=""
            THRUST=""
            POWER=""
            ASSUMEDTEMP=""
            V1=""
            VR=""
            V2=""
            Vref=""
            EOGO=""
            EOSP=""
            AEGO=""
            AIRPORTELEV=""
            EOSID=""
            TODA=""
            ASDA=""
            CG=""  
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
                    
                    if "TOW" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            TOW=splittedbody[items+2]
                except:
                    TOW=""
                try:
                    if "MFRH" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            MFRH=splittedbody[items+2]
                except:
                    MFRH=""
                try:
                    if "MLOH" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            MLOH=splittedbody[items+2]
                except:
                    MLOH=""
                try:
                    if "FLAPS" in splittedbody[items]:
                        if splittedbody[items+1]==":":
                            FLAPS=splittedbody[items+2]
                except:
                    FLAPS=""
                try:
                        
                    if "THRUST" in splittedbody[items]:
                        if splittedbody[items+1]=="RATING":
                            THRUST=splittedbody[items+3]
                except:
                    THRUST=""
                try:
                    if "POWER" in splittedbody[items]:
                        if splittedbody[items+1]=="SETTING":
                            POWER=splittedbody[items+3]
                            ASSUMEDTEMP=splittedbody[items+4]
                except:
                    ASSUMEDTEMP=""
                try:
                    if "V1" in splittedbody[items]:
                        V1=splittedbody[items+1]
                except:
                    V1=""
                try:
                    if "VR" in splittedbody[items]:
                        VR=splittedbody[items+1]
                except:
                    VR=""
                try:    
                    if "V2" in splittedbody[items]:
                        V2=splittedbody[items+1]
                except:
                    V2=""
                try:
                    if "Vref" in splittedbody[items]:
                        Vref=splittedbody[items+1]
                except:
                    Vref=""
                try:
                    if "EOGO" in splittedbody[items]:
                        EOGO=splittedbody[items+1]
                
                    if "EO" in splittedbody[items]:
                        if "GO" in splittedbody[items+1]:
                            EOGO=splittedbody[items+2]
                except:
                    EOGO=""
                try:
                        
                    if "EOSP" in splittedbody[items]:
                        EOSP=splittedbody[items+1]
                    if "EO" in splittedbody[items]:
                        if "STOP" in splittedbody[items+1]:
                            EOSP=splittedbody[items+2]
                except:
                    EOSP=""
                try:
                    if "AEGO" in splittedbody[items]:
                        AEGO=splittedbody[items+1]
                    if "ALL" in splittedbody[items]:
                        if "GO" in splittedbody[items+1]:
                            AEGO=splittedbody[items+2]
                except:
                    AEGO=""
                try:
                    
                    if "AIRPORT" in splittedbody[items]:
                        if "ELEV" in splittedbody[items+1]:
                            AIRPORTELEV=splittedbody[items+2]
                except:
                    AIRPORTELEV=""
                try:
                    if "ENGINE" in splittedbody[items]:
                        if "FAILURE" in splittedbody[items+1]:
                            
                            for cntEOSID in range(1000):
                                if 'WEIGHT' in splittedbody[items+3+cntEOSID]:
                                    break
                                else:
                                    EOSID=EOSID+" "+splittedbody[items+3+cntEOSID]
                except:
                    EOSID=""
                try:
                    if "TODA" in splittedbody[items]:
                        TODA=splittedbody[items+1]
                except:
                    TODA=""
                try:
                    if "ASDA" in splittedbody[items]:
                        ASDA=splittedbody[items+1]
                except:
                    ASDA=""
                try:
                    if "TAKEOFF" in splittedbody[items]:
                        CG=splittedbody[items+2]
                except:
                    CG=""
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
            CONDITION,
            WIND,
            OAT,
            QNH,
            PACKS,
            TOW,
            MFHR,
            MLOH,
            FLAPS,
            THRUST,
            POWER,
            ASSUMEDTEMP,
            V1,
            VR,
            V2,
            Vref,
            EOGO,
            EOSP,
            AEGO,
            AIRPORTELEV,
            EOSID,
            TODA,
            ASDA,
            CG,  
            AIRPORTDB,
            POLICYDB,
            MELCDL)
            print(listtowrite)
            try:
                cur=conn.cursor()
                cur.execute("INSERT INTO dbo.[OPT Takeoff] VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                listtowrite)
                conn.commit()
            except:
                pass
    try:
        messages[cntm+1].Move(ArchieveTO)
    except:
        pass
            
