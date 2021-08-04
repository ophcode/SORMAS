import csv
import datetime
import sys
import os
import tkinter
import re
import shutil

from mailmerge import MailMerge #Externe Library für docx-Serienbriefe
from docx2pdf import convert    #Externe Library für pdf-Erstellung mit Word
import win32com.client          #Externe Library für Email-Versand per Outlook

import tkinter as tk
import tkinter.filedialog as fd
import tkinter.simpledialog as sd

def send_mail(email,subject,body,attachment_path_list):
    outlook = win32com.client.Dispatch("outlook.application")
    msg = outlook.CreateItem(0)
    msg.To = email
    msg.Subject = subject
    msg.HTMLbody = body
    msg.SentOnBehalfOfName = "x@x.de" #Getestet, funktioniert falls Zugriff auf das Konto besteht
    for attachment_path in attachment_path_list:
        att = os.path.abspath(attachment_path)
        msg.Attachments.Add(att)
    msg.display()   #Email anzeigen, kein automatisches senden #msg.Send()  
   
def csv_to_dict(filename):  #Convert Sormas csv export to nested dictionary, first order key: person Id, second order keys specified in file header
    with open(filename,encoding="UTF-8",newline="") as csvfile:
        sormasreader = csv.reader(csvfile, delimiter=";", quotechar='"')
        d = {}
        next(sormasreader)
        header_keys = next(sormasreader)
        for row in sormasreader:
            dd=dict(zip(header_keys,row))
            d[row[0]]=dd
        return d 

def read_folder(folderpath):
    d={}
    for file in os.listdir(folderpath):
        print(file)
        if file.endswith(".csv"):
            d[file[:-4]]=csv_to_dict(os.path.join(folderpath,file))
    return d

def convert_date(datestring): #Return DD.MM.YYYY if String is YYYY-MM-DD (HH:MM:SS.xxxxxx), return String otherwise.
    if re.match(r"[0-9][0-9][0-9][0-9]\-[0-9][0-9]\-[0-9][0-9] *",datestring) and len(datestring)<30:
        return datestring[8:10]+"."+datestring[5:7]+"."+datestring[:4]
    return datestring

def date_or_empty(datestring): #Convert timestamp to date, default to empty for invalid/empty data
    try:
        datum = datetime.datetime.strptime(datestring, "%Y-%m-%d %H:%M:%S")
    except:
        return ""
    return datum.strftime("%d.%m.%Y")

def contacts_from(quarantinefrom):
    #Defaults to empty in case of empty/invalid
    try:
        quarantinestart = datetime.datetime.strptime(quarantinefrom, "%Y-%m-%d %H:%M:%S")
    except:
        return ""
    contactsfrom_date=(quarantinestart-datetime.timedelta(days=2))
    return contactsfrom_date.strftime("%d.%m.%Y")

def is_adult(birthdate): #Date format: yyyy-mm-dd
    #Defaults to 'True' in case of missing or malformated data
    today = datetime.date.today()
    try:
        person_birthdate=datetime.date.fromisoformat(birthdate)
    except:
        return True
    if (today-person_birthdate).days > 6574:
        return True
    return False

class SORMAS:
    def __init__(self,input_dir,outputfolder,startdatetime,enddatetime,Sachbearbeiter):  #Directory of Sormas dump
        self.enddatetime = enddatetime
        self.startdatetime = startdatetime   
        self.S = read_folder(input_dir)
        self.msglog=""
        self.Sachbearbeiter=Sachbearbeiter
        self.inputdocxpath=os.path.join("Vorlagen")
        self.outputfolder=outputfolder
        self.c_id_list=[] #Who shows up in list
        self.not_executable_tasks=[] #Who shows up in list as red
        self.selection_mail=[] #Who gets the mail
        self.selection_mail_ne=[] #Who gets the "nicht erreicht"-letter
        self.notificationdict={} #Caseid -> log
        self.mail_sent=[]
        self.letter_sent=[] #"Nicht erreicht" only
        self.removed_tasks=[] #Aufgabe als entfernt markiert, aber Deckblatt wird erstellt
    
    def get_mail(self, c_id, context = "cases"):
        l = [d for d in self.S["person_contact_details"].values() if d["person_id"] == self.S[context][c_id]["person_id"]]
        if len([d for d in l if d["personcontactdetailtype"] == "EMAIL" and d["primarycontact"] == "t"])==1:
            return [d["contactinformation"] for d in l if d["personcontactdetailtype"] == "EMAIL" and d["primarycontact"] == "t"][0]
        elif len([d for d in l if d["personcontactdetailtype"] == "EMAIL"])>0:
            return [d["contactinformation"] for d in l if d["personcontactdetailtype"] == "EMAIL"][0]
        else:
            if context=="cases":
                self.note(c_id,"✉ Keine Email angegeben")
            return ""  
    
    def note(self, c_id, msg):
        p=self.S["persons"][self.S["cases"][c_id]["person_id"]]
        p=self.S["persons"][self.S["cases"][c_id]["person_id"]]
        print("HINWEIS: "+p["lastname"]+", "+p["firstname"]+" - "+msg)
        self.msglog+="HINWEIS: "+p["lastname"]+", "+p["firstname"]+" - "+msg+"\n"
        self.notificationdict[c_id]+=msg+"\n"
    
    def send_standard_mail(self, c_id, email):
        subject="Quarantäne-Anschreiben & Kontaktpersonenermittlung"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation Kontaktpersonen.pdf"),os.path.join("Vorlagen","Merkblatt für die häusliche Isolation positiv Getestete.pdf")]
        p=self.S["persons"][self.S["cases"][c_id]["person_id"]]
        if not is_adult(p["birthdate_yyyy"]+"-"+p["birthdate_mm"].zfill(2)+"-"+p["birthdate_dd"].zfill(2)):
            with open(os.path.join("Vorlagen","Email_u18.txt"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email.txt"),encoding="utf-8") as f:
                body_template = f.read()
        c=self.S["cases"][c_id]
        Anrede="Sehr geehrte*r"
        if p["sex"]=="MALE":
            Anrede="Sehr geehrter Herr"
        if p["sex"]=="FEMALE":
            Anrede="Sehr geehrte Frau"
        attachment_path_list=attachment_paths.copy()
        shutil.copy(os.path.join("Vorlagen","Kontaktpersonen_Nachname_Vorname.xlsx"),os.path.join(self.outputfolder,"pdf"))
        if not os.path.exists(os.path.join(self.outputfolder,"pdf","Kontaktpersonen_"+p["lastname"]+"_"+p["firstname"]+".xlsx")):
            os.rename(os.path.join(self.outputfolder,"pdf","Kontaktpersonen_Nachname_Vorname.xlsx"), os.path.join(self.outputfolder,"pdf","Kontaktpersonen_"+p["lastname"]+"_"+p["firstname"]+".xlsx"))
        attachment_path_list.append(os.path.join(self.outputfolder,"pdf",p["lastname"]+"-"+p["firstname"]+"-"+c["externaltoken"][-5:]+".pdf"))
        attachment_path_list.append(os.path.join(self.outputfolder,"pdf","Kontaktpersonen_"+p["lastname"]+"_"+p["firstname"]+".xlsx"))
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstname"]).replace("{Nachname}",p["lastname"]).replace("{Kontakte_ab}",contacts_from(c["quarantinefrom"])).replace("{Sachbearbeiter}",self.Sachbearbeiter)
        send_mail(email,subject,body,attachment_path_list)  

    def send_mail_no_contacts(self, c_id, email):
        subject="Quarantäne-Anschreiben & Kontaktpersonenermittlung"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation positiv Getestete.pdf")]
        p=self.S["persons"][self.S["cases"][c_id]["person_id"]]
        if not is_adult(p["birthdate_yyyy"]+"-"+p["birthdate_mm"].zfill(2)+"-"+p["birthdate_dd"].zfill(2)):
            with open(os.path.join("Vorlagen","Email_ohne_KP_u18.txt"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_ohne_KP.txt"),encoding="utf-8") as f:
                body_template = f.read()
        c=self.S["cases"][c_id]
        Anrede="Sehr geehrte*r"
        if p["sex"]=="MALE":
            Anrede="Sehr geehrter Herr"
        if p["sex"]=="FEMALE":
            Anrede="Sehr geehrte Frau"
        attachment_path_list=attachment_paths.copy()
        attachment_path_list.append(os.path.join(self.outputfolder,"pdf",p["lastname"]+"-"+p["firstname"]+"-"+c["externaltoken"][-5:]+".pdf"))
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstname"]).replace("{Nachname}",p["lastname"]).replace("{Kontakte_ab}",contacts_from(c["quarantinefrom"])).replace("{Sachbearbeiter}",self.Sachbearbeiter)
        send_mail(email,subject,body,attachment_path_list)  

    def send_contact_mail(self, c_id, email):
        subject="Isolationsschreiben"
        body_template=""
        attachment_paths=[os.path.join("Vorlagen","Merkblatt für die häusliche Isolation Kontaktpersonen.pdf")]
        p=self.S["persons"][self.S["contacts"][c_id]["person_id"]]
        if not is_adult(p["birthdate_yyyy"]+"-"+p["birthdate_mm"].zfill(2)+"-"+p["birthdate_dd"].zfill(2)):
            with open(os.path.join("Vorlagen","Email_KP_u18.txt"),encoding="utf-8") as f:
                body_template = f.read()
        else:
            with open(os.path.join("Vorlagen","Email_KP.txt"),encoding="utf-8") as f:
                body_template = f.read()
        c=self.S["contacts"][c_id]
        Anrede="Sehr geehrte*r"
        if p["sex"]=="MALE":
            Anrede="Sehr geehrter Herr"
        if p["sex"]=="FEMALE":
            Anrede="Sehr geehrte Frau"
        attachment_path_list=attachment_paths.copy()
        aktenzeichen=c.get("externaltoken","")
        if aktenzeichen=="":
            aktenzeichen=c["uuid"].split("-")[0]
        print(os.path.join(self.outputfolder,"pdf",p["lastname"]+"-"+p["firstname"]+"-"+aktenzeichen+".pdf"))
        attachment_path_list.append(os.path.join(self.outputfolder,"pdf",p["lastname"]+"-"+p["firstname"]+"-"+aktenzeichen+".pdf"))
        body=body_template.replace("{Anrede}",Anrede).replace("{Vorname}",p["firstname"]).replace("{Nachname}",p["lastname"]).replace("{Sachbearbeiter}",self.Sachbearbeiter)
        send_mail(email,subject,body,attachment_path_list)  
        
    def create_docx(self, c_id, docx_file, suffix="", context="cases"):
        if not os.path.exists(os.path.join(self.outputfolder,"pdf")):
            os.makedirs(os.path.join(self.outputfolder,"pdf"))
        d={}
        c=self.S[context][c_id]
        p=self.S["persons"][self.S[context][c_id]["person_id"]]
        l=self.S["locations"][self.S["persons"][self.S[context][c_id]["person_id"]]["address_id"]]
        Anrede = "Sehr geehrte/r"
        if p["sex"]=="MALE":
            Anrede="Sehr geehrter Herr"
        elif p["sex"]=="FEMALE":
            Anrede="Sehr geehrte Frau"
        else:
            self.note(c_id, "🚻 Geschlecht nicht ausgefüllt, benutze neutrale Anrede für Mail/Bescheid")
        d["Anrede"]=Anrede
        d["Sachbearbeiter"]=self.Sachbearbeiter
        d["externaltoken"]=c.get("externaltoken","")
        aktenzeichen=""
        if d["externaltoken"]=="":
            aktenzeichen=c["uuid"]
            d["externaltoken"]=c["uuid"]
        else:
            aktenzeichen=d["externaltoken"][-5:]
        d["firstname"]=p["firstname"]
        d["lastname"]=p["lastname"]
        d["street"]=l["street"]
        d["housenumber"]=l["housenumber"]
        d["postalcode"]=l["postalcode"]
        d["quarantinefrom"]=date_or_empty(c["quarantinefrom"])
        d["quarantineto"]=date_or_empty(c["quarantineto"])
        d["PCRSatz"]=""
        
        if context=="cases" and self.get_earliest_positive_PCR_date(c_id)!="":
            d["PCRSatz"]=" Der Nachweis darüber erfolgte mit Abstrich vom "+date_or_empty(self.get_earliest_positive_PCR_date(c_id))+" mittels PCR-Test."
        if not is_adult(p["birthdate_yyyy"]+"-"+p["birthdate_mm"].zfill(2)+"-"+p["birthdate_dd"].zfill(2)):
            if os.path.exists(os.path.join(self.inputdocxpath,docx_file+"_u18.docx")):
                self.fill_file(os.path.join(self.inputdocxpath,docx_file+"_u18.docx"),os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+aktenzeichen.split("-")[0]+suffix+".docx"),d)
            else:
                self.fill_file(os.path.join(self.inputdocxpath,docx_file+".docx"),os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+aktenzeichen.split("-")[0]+suffix+".docx"),d)
        else:
            self.fill_file(os.path.join(self.inputdocxpath,docx_file+".docx"),os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+aktenzeichen.split("-")[0]+suffix+".docx"),d)
        convert(os.path.join(self.outputfolder,"pdf",d["lastname"]+"-"+d["firstname"]+"-"+aktenzeichen.split("-")[0]+suffix+".docx"))

    def fill_file(self,input_docx_path,output_docx_path,fielddict): #Serienbrief erstellen, keys in fielddict müssen keys in docx entsprechen
        with MailMerge(input_docx_path) as document:
            document.merge_templates([fielddict],separator="page_break")
            document.write(output_docx_path)
 
    def tasks_completed_on_date(self,startdatetime,enddatetime,taskstatus="DONE",tasktype="CASE_INVESTIGATION"): 
        #Output list of task ids with investigation tasks marked as taskstatus on specified date. Date format "YYYY-MM-DD" as String
        #Taskstatus: "DONE", "PENDING", "NOT_EXECUTABLE", "REMOVED"
        completed=[]
        for key in self.S["tasks"]:
            if (self.S["tasks"][key]["tasktype"]==tasktype) & (self.S["tasks"][key]["taskstatus"]==taskstatus) & (self.S["tasks"][key]["statuschangedate"]>=startdatetime) & (self.S["tasks"][key]["statuschangedate"]<=enddatetime) :
                completed.append(key)
        return completed

    def append_reimport_csv(self, c_id, note=""):
        with open(os.path.join(self.outputfolder, "mail_reimport.csv"),"a",encoding="UTF-8") as output_file:
                c=self.S["cases"][c_id]
                p=self.S["persons"][self.S["cases"][c_id]["person_id"]]
                if note == "":
                    output_file.write(";".join([c["disease"],p["firstname"],p["lastname"],p["sex"],p["birthdate_dd"],p["birthdate_mm"],p["birthdate_yyyy"],"Berlin","SK Berlin Friedrichshain-Kreuzberg","FK","","NO_FACILITY","",datetime.date.today().strftime("%d.%m.%Y"),"TRUE",""])+"\n")
                else:
                    output_file.write(";".join([c["disease"],p["firstname"],p["lastname"],p["sex"],p["birthdate_dd"],p["birthdate_mm"],p["birthdate_yyyy"],"Berlin","SK Berlin Friedrichshain-Kreuzberg","FK","","NO_FACILITY","",datetime.date.today().strftime("%d.%m.%Y"),"TRUE",note+c["additionaldetails"]])+"\n")

    def get_earliest_positive_sample(self,c_id): #Output first positive test for c_id (Output=dict) merged with sample_test entry belonging to sample
        sample_list=[self.S["samples"][key] for key in self.S["samples"].keys() if self.S["samples"][key]["associatedcase_id"]==c_id]
        if len(sample_list)==0:
            return dict([(key,"") for key in next(iter(self.S["sample_tests"].values())).keys()]) | dict([(key,"") for key in next(iter(self.S["samples"].values())).keys()])
        else:
            firstsample = min(sample_list,key=lambda d: d["sampledatetime"])
            sampletest_list=[self.S["sample_tests"][key] for key in self.S["sample_tests"].keys() if self.S["sample_tests"][key]["sample_id"]==firstsample["id"]]
            if len(sampletest_list)==0:
                self.note(c_id,"Kein Test angelegt")
                return firstsample | dict([(key,"") for key in next(iter(self.S["samples"].values())).keys()])
            else:
                return firstsample | sampletest_list[0]
    
    def get_earliest_positive_PCR_date(self,c_id): #Output first positive PCR for c_id (Output=dict) merged with sample_test entry belonging to sample
        sample_list=[self.S["samples"][key] for key in self.S["samples"].keys() if self.S["samples"][key]["associatedcase_id"]==c_id]
        if len(sample_list)==0:
            return ""
        else:
            firstsample = min(sample_list,key=lambda d: d["sampledatetime"])
            sampletest_list=[self.S["sample_tests"][key] for key in self.S["sample_tests"].keys() if self.S["sample_tests"][key]["sample_id"]==firstsample["id"] and self.S["sample_tests"][key]["testtype"]=="PCR_RT_PCR"]
            if len(sampletest_list)==0:
                self.note(c_id,"❗ Kein PCR vorhanden")
                return ""
            else:
                return firstsample.get("sampledatetime")
    

def initialize(startdate, enddate, input_dir,Sachbearbeiter,outputfolder):
    startdatetime=startdate[6:10]+"-"+startdate[3:5]+"-"+startdate[:2]+startdate[10:]
    enddatetime=enddate[6:10]+"-"+enddate[3:5]+"-"+enddate[:2]+enddate[10:]
    
    S = SORMAS(input_dir, outputfolder, startdatetime, enddatetime, Sachbearbeiter)
    
    print("\nErledigte Aufgaben suchen:\n")
    completed_task_keys=S.tasks_completed_on_date(startdatetime, enddatetime, taskstatus="DONE")
    completed_task_ids=[]
        
    for key in completed_task_keys:
        c_id=S.S["tasks"][key]["caze_id"]
        completed_task_ids.append(c_id)
        S.notificationdict[c_id]=""
        S.note(c_id,"Aufgabenkommentar: "+S.S["tasks"][key]["creatorcomment"])
        S.note(c_id,"Ausführungskommentar: "+S.S["tasks"][key]["assigneereply"])
    S.c_id_list=completed_task_ids
    
    not_executable_task_keys=S.tasks_completed_on_date(startdatetime, enddatetime, taskstatus="NOT_EXECUTABLE")
    not_executable_task_ids=[]

    for key in not_executable_task_keys:
        c_id=S.S["tasks"][key]["caze_id"]
        not_executable_task_ids.append(c_id)
        S.notificationdict[c_id]=""
        S.note(c_id,"NICHT AUSFÜHRBAR: Aufgabenkommentar: "+S.S["tasks"][key]["creatorcomment"])
        S.note(c_id,"Ausführungskommentar: "+S.S["tasks"][key]["assigneereply"])
    S.not_executable_tasks=not_executable_task_ids

    print("\nAusgabe in Ordner: "+S.outputfolder)
    if S.outputfolder=="":
        S.outputfolder=datetime.datetime.today().strftime("%Y-%m-%d %H-%M-%S")
    if not os.path.exists(S.outputfolder):
        os.makedirs(S.outputfolder)
    with open(os.path.join(S.outputfolder, "mail_reimport.csv"),"w",encoding="UTF-8") as output_file:
        output_file.write("CaseData;Person;Person;Person;Person;Person;Person;CaseData;CaseData;CaseData;CaseData;CaseData;CaseData;CaseData;CaseData;CaseData\n")
        output_file.write("disease;person.firstName;person.lastName;person.sex;person.birthdateDD;person.birthdateMM;person.birthdateYYYY;region;district;community;facilityType;healthFacility;healthFacilityDetails;quarantineOrderedOfficialDocumentDate;quarantineOrderedOfficialDocument;additionalDetails\n")
    return S
    
class Application(tk.Frame): #GUI

    def __init__(self, master=None):
        super().__init__(master)
        self.buttonfont=('Berlin Type Office Regular',18)
        self.bw= 30 # buttwonwidth
        self.bh= 2  # buttonheight
        self.bc1= "white"  # button color (unactived)
        self.bc2= "ghost white"  # button color (activated)
        self.inp_dir=""
        self.startdate=""
        self.enddate=""
        self.Sachbearbeiter=""
        self.opd=""
        self.S = ""
        self.lb = ""
        self.master = master
        self.pack()
        self.create_widgets()
        self.mail_win = ""
        self.list_win = ""
        self.mail_answer = ""
        
    def create_widgets(self):
        self.input_dir = tk.Button(self, text = " 1. Sormas-Export-Ordner wählen", command=self.choose_dir, anchor="w", font=self.buttonfont, width=self.bw, height=self.bh, bg=self.bc1)
        self.input_dir.pack(side="top")
        self.input_startdatetime = tk.Button(self,text= " 2. Startzeitpunkt wählen", command=self.choose_startdate, anchor="w", font=self.buttonfont, width=self.bw, height=self.bh, bg=self.bc1)
        self.input_startdatetime.pack(side="top")
        self.input_enddatetime = tk.Button(self, text=" 3. Endzeitpunkt wählen", command=self.choose_enddate, anchor="w", font=self.buttonfont, width=self.bw, height=self.bh, bg=self.bc1)
        self.input_enddatetime.pack(side="top")
        self.input_sachbearbeiter = tk.Button(self, text=" 4. Sachbearbeiter*in", command=self.choose_sachbearbeiter, anchor="w", font=self.buttonfont,width=self.bw, height=self.bh, bg=self.bc1)
        self.input_sachbearbeiter.pack(side="top")
        #self.opd = tk.Button(self, text="5. Output - Verzeichnis wählen (optional)", command=self.choose_opd)
        #self.opd.pack(side="top")
        self.run = tk.Button(self, text=" 5. Starten", command=self.run_script, anchor="w",  font=self.buttonfont, width=self.bw, bg='white')
        self.run.pack(side="top")
        self.run_contacts = tk.Button(self, text=" 6. KP-Mails", command=self.send_contact_mail, anchor="w",  font=self.buttonfont, width=self.bw, bg='white')
        self.run_contacts.pack(side="top")
        self.quit = tk.Button(self, text=" BEENDEN", command=self.master.destroy, anchor="w", font=self.buttonfont, width=self.bw, bg='misty rose')
        self.quit.pack(side="bottom")
    
    def send_mail(self):
        self.S.selection_mail = [self.S.c_id_list[x] for x in self.lb.curselection() if x < len(self.S.c_id_list)]
        self.S.selection_mail_ne = [self.S.not_executable_tasks[x-len(self.S.c_id_list)] for x in self.lb.curselection() if x >= len(self.S.c_id_list)]
        for c_id in self.S.selection_mail:
            mail=self.S.get_mail(c_id)
            self.send_one_mail(c_id, mail)
        for c_id in self.S.selection_mail_ne:
            text=self.S.notificationdict[c_id]
            answer = tk.messagebox.askyesno("","Nicht-erreicht-Brief drucken?\n"+text)
            if answer:
                self.S.create_docx(c_id,"Anschreiben nicht erreichte Indices", suffix = "_ne")
                os.startfile(os.path.join(self.S.outputfolder,"pdf",self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["lastname"]+"-"+self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["firstname"]+"-"+self.S.S["cases"][c_id]["externaltoken"][-5:]+"_ne.pdf"))
                self.S.append_reimport_csv(c_id, note="Nicht erreicht Brief am "+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt")

    def send_contact_mail(self):
        startdatetime=self.startdate[6:10]+"-"+self.startdate[3:5]+"-"+self.startdate[:2]+self.startdate[10:]
        enddatetime=self.enddate[6:10]+"-"+self.enddate[3:5]+"-"+self.enddate[:2]+self.enddate[10:]
        if self.S == "":
            self.S = SORMAS(self.inp_dir, self.opd, startdatetime, enddatetime, self.Sachbearbeiter)
        
        completed_task_keys=[self.S.S["tasks"][key]["contact_id"] for key in  self.S.tasks_completed_on_date(startdatetime, enddatetime, taskstatus="DONE", tasktype = "CONTACT_INVESTIGATION")]
        not_executable_task_ids=[self.S.S["tasks"][key]["contact_id"] for key in  self.S.tasks_completed_on_date(startdatetime, enddatetime, taskstatus="NOT_EXECUTABLE", tasktype = "CONTACT_INVESTIGATION")]        
        self.S.outputfolder=datetime.datetime.today().strftime("%Y-%m-%d %H-%M-%S")+"_contacts"
        if not os.path.exists(self.S.outputfolder):
            os.makedirs(self.S.outputfolder)
        with open(os.path.join(self.S.outputfolder, "mail_reimport.csv"),"w",encoding="UTF-8") as output_file:
            output_file.write("disease;person.firstName;person.lastName;person.sex;person.birthdateDD;person.birthdateMM;person.birthdateYYYY;quarantineOrderedOfficialDocument;quarantineOrderedOfficialDocumentDate\n") #TODO region / district (Default for FK)
            for c_id in completed_task_keys:
                if self.S.S["contacts"][c_id]["contactclassification"]!="NO_CONTACT":
                    self.S.create_docx(c_id, "Isolationsbescheinigung", context="contacts")
                    email=self.S.get_mail(c_id, context = "contacts")
                    self.S.send_contact_mail(c_id, email)
                    c=self.S.S["contacts"][c_id]
                    p=self.S.S["persons"][self.S.S["contacts"][c_id]["person_id"]]
                    output_file.write(";".join([c["disease"],p["firstname"],p["lastname"],p["sex"],p["birthdate_dd"],p["birthdate_mm"],p["birthdate_yyyy"],"TRUE",datetime.date.today().strftime("%d.%m.%Y")])+"\n")
            for c_id in not_executable_task_ids:
                print(c_id)
                self.S.create_docx(c_id, "KP_ne", context="contacts")
                email=self.S.get_mail(c_id, context = "contacts")
                self.S.send_contact_mail(c_id, email)
                c=self.S.S["contacts"][c_id]
                p=self.S.S["persons"][self.S.S["contacts"][c_id]["person_id"]]
                output_file.write(";".join([c["disease"],p["firstname"],p["lastname"],p["sex"],p["birthdate_dd"],p["birthdate_mm"],p["birthdate_yyyy"],"TRUE",datetime.date.today().strftime("%d.%m.%Y")])+"\n")
    
  
    def send_one_mail(self, c_id, mail):
        self.mail_window(c_id)
        self.list_win.wait_window(self.mail_win)
        print(self.mail_answer)
        if self.mail_answer == 1:
            self.S.create_docx(c_id,"Anschreiben Indices")
            self.S.send_standard_mail(c_id, mail)
            self.S.append_reimport_csv(c_id)
            #TODO abgeschlossener Q-Zeitraum?
        elif self.mail_answer == 2:
            self.S.create_docx(c_id,"Anschreiben Indices")
            self.S.send_mail_no_contacts(c_id, mail)
            self.S.append_reimport_csv(c_id)
        elif self.mail_answer == 3:
            self.S.create_docx(c_id,"Anschreiben Indices")
            os.startfile(os.path.join(self.S.outputfolder,"pdf",self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["lastname"]+"-"+self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["firstname"]+"-"+self.S.S["cases"][c_id]["externaltoken"][-5:]+".pdf"))
            self.S.append_reimport_csv(c_id, note="Bescheid per Brief am "+datetime.datetime.today().strftime("%d.%m.%Y")+" verschickt")

            
    def select_all(self):
        self.lb.select_set(0, tk.END)
        
    def listbox(self):
        self.list_win = tk.Toplevel(self)
        self.list_win.title("Mails "+self.S.startdatetime+" - "+self.S.enddatetime)
        self.lb=tk.Listbox(self.list_win, height=min(len(self.S.c_id_list+self.S.not_executable_tasks),28), width=30, font=font, selectmode="multiple")
        for i,c_id in enumerate(self.S.c_id_list+self.S.not_executable_tasks):
            casestr = self.S.notificationdict[c_id].split("Aufgabenkommentar:")[1][:4]+" \t"+self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["firstname"]+" "+self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["lastname"]
            self.lb.insert(i,casestr)
            if i>=len(self.S.c_id_list):
                self.lb.itemconfig(i, {'bg':'light pink'})
        self.lb.grid(row=0, column=0, columnspan=2, sticky=tk.N)
        b1 = tk.Button(self.list_win, text="Mail schicken", command=self.send_mail, font=self.buttonfont)
        b1.grid(row=1, column=1, sticky=tk.N)
        b2 = tk.Button(self.list_win, text="Alle auswählen", command=self.select_all, font=self.buttonfont)
        b2.grid(row=1, column=0, sticky=tk.N)
    
        
    def mail_window(self, c_id):
        def return_value(value):
            self.mail_answer = value
            self.mail_win.destroy()
        self.mail_win = tk.Toplevel(self, bg = "white")
        if self.S.S["cases"][c_id]["quarantine"]=="INSTITUTIONELL":
            self.S.note(c_id, "🚑 Quarantäneort als 'institutionell' markiert ("+self.S.S["cases"][c_id]["healthfacilitydetails"]+"), prüfen ob Brief/Mail verschickt werden muss.\n")
        if self.S.S["cases"][c_id]["quarantineorderedofficialdocument"]=="t":
            self.S.note(c_id, "⚠ Anordnung wurde bereits am "+self.S.S["cases"][c_id]["quarantineorderedofficialdocumentdate"]+" als verschickt markiert, bitte prüfen ob Bescheid erneut versendet werden soll.\n")
        if self.S.S["cases"][c_id]["caseclassification"]=="NO_CASE":
            self.S.note(c_id,"❗ Akte wurde als 'KEIN FALL' markiert")
        #if d["street"]=="" or d["postalcode"]=="":
        #    self.note(c_id, "⚠ Adresse unvollständig")
        if self.S.S["cases"][c_id]["quarantinefrom"]=="" or self.S.S["cases"][c_id]["quarantineto"]=="":
            self.S.note(c_id, "⚠ Quarantänezeitraum unvollständig")
        self.mail_win.title(self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["firstname"]+" "+self.S.S["persons"][self.S.S["cases"][c_id]["person_id"]]["lastname"])
        text = "Hinweise:\n"+self.S.notificationdict[c_id]
        text+= "\n Quarantänezeitraum: "+date_or_empty(self.S.S["cases"][c_id]["quarantinefrom"])+" bis "+date_or_empty(self.S.S["cases"][c_id]["quarantineto"])+"\n"
        text+= "Symptombeginn: "+date_or_empty(self.S.S["case_symptoms"][self.S.S["cases"][c_id]["symptoms_id"]]["onsetdate"])+"  Symptomende: "+date_or_empty(self.S.S["cases"][c_id]["outcomedate"]+"\n\n")
        text+= self.S.S["cases"][c_id]["additionaldetails"]
        label = tk.Label(self.mail_win, text=text, bg = "white", font=font, wraplength=1000).grid(row=0, column=0, columnspan=4, sticky=tk.N)
        b1 = tk.Button(self.mail_win, text="Mail + KP", command = lambda: return_value(1), font=self.buttonfont).grid(row=1, column=0, sticky=tk.N)
        b2 = tk.Button(self.mail_win, text="Nur Mail", command = lambda: return_value(2), font=self.buttonfont).grid(row=1, column=1, sticky=tk.N)
        b3 = tk.Button(self.mail_win, text="Brief", command = lambda: return_value(3), font=self.buttonfont).grid(row=1, column=2, sticky=tk.N)
        b4 = tk.Button(self.mail_win, text="Nichts", command =  lambda : return_value(0), font=self.buttonfont).grid(row=1, column=3, sticky=tk.N)
  
    def choose_dir(self):
        self.inp_dir = fd.askdirectory().replace("/","\\")
        self.input_dir["bg"]=self.bc2
        self.input_dir["text"]="1. Sormas-Export-Ordner wählen\n"+self.inp_dir.split("\\")[-1]
        print("Input-Ordner: "+self.inp_dir)      

    def choose_startdate(self):
        self.startdate = sd.askstring("Datum", "Ermittlungsdatum eingeben (TT.MM.JJJJ HH:MM)", parent=self.master)
        self.input_startdatetime["bg"]=self.bc2
        self.input_startdatetime["text"]="2. Startzeitpunkt wählen\n"+self.startdate
        print("Datum: "+self.startdate)
    
    def choose_enddate(self):
        self.enddate = sd.askstring("Datum", "Ermittlungsdatum eingeben (TT.MM.JJJJ HH:MM)", parent=self.master)
        self.input_enddatetime["bg"]=self.bc2
        self.input_enddatetime["text"]="3. Endzeitpunkt wählen\n"+self.enddate
        print("Datum: "+self.enddate)
        
    def choose_sachbearbeiter(self):
        self.Sachbearbeiter = sd.askstring("Sachbearbeiter*in", "Stellenzeichen eingeben", parent=self.master)
        self.input_sachbearbeiter["bg"]=self.bc2
        self.input_sachbearbeiter["text"]="4. Sachbearbeiter*in\n"+self.Sachbearbeiter
        print("Sachbearbeiter: "+self.Sachbearbeiter)
        
    def choose_opd(self):
        self.opd = fd.askdirectory().replace("/","\\")
        print("Output-Ordner: "+self.opd)

    def run_script(self):
        print("Initialisierung:")
        self.S=initialize(self.startdate, self.enddate, self.inp_dir, self.Sachbearbeiter, self.opd)
        self.listbox()
        

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(master=root)
    icon = tk.PhotoImage(file = os.path.join("Vorlagen","icon.png"))
    root.iconphoto(False, icon)
    global font
    font=('Berlin Type Office Regular',14)
    app.mainloop()