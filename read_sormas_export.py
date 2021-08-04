"""
Sammlung von Funktionen, um den SORMAS-Export auszulesen und damit zu arbeiten.
Funktionsweise:
SORMAS -> Statistik -> Datenbank Export-> Alle SORMAS-Daten auswählen -> Entpacken, in "sormas_export" umbenennen und in Ordner von Skript packen. Dann Skript starten.
Erstellt ein Dictionary (S), in dem alle tabellen gelistet sind
S["cases"] enthält alle Fälle,
S["cases"]["1"] enthält Infos zum Fall mit der ID 1
S["cases"]["1"]["reportdate"] ist das Meldedatum des Falles mit ID 1
S["persons"][S["cases"]["1"][person_id]] enthält Daten zur Person, die zum Fall mit ID 1 gehört...
S["persons"][S["cases"]["1"][person_id]]["sex"] ...hier zum Beispiel das Geschlecht.
"""

import csv
import datetime
import os
import re

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

def read_folder(folderpath): # Read all csv files in SORMAS output directory
    d={}
    for file in os.listdir(folderpath):
        print(file)
        if file.endswith(".csv"):
            d[file[:-4]]=csv_to_dict(os.path.join(folderpath,file))
            print(str(len(d[file[:-4]])+" Einträge")
    return d

def convert_date(datestring): #Return DD.MM.YYYY if String is YYYY-MM-DD (HH:MM:SS.xxxxxx), return input String otherwise.
    if re.match(r"[0-9][0-9][0-9][0-9]\-[0-9][0-9]\-[0-9][0-9] *",datestring) and len(datestring)<30:
        return datestring[8:10]+"."+datestring[5:7]+"."+datestring[:4]
    return datestring

def date_or_empty(datestring): #Convert timestamp to date, default to empty for invalid/empty data
    try:
        datum = datetime.datetime.strptime(datestring, "%Y-%m-%d %H:%M:%S")
    except:
        return ""
    return datum.strftime("%d.%m.%Y")

def is_adult(birthdate): #Date format: yyyy-mm-dd
    #Defaults to 'True' in case of missing or malformated data
    #Is a bit lazy but works in mostg cases
    today = datetime.date.today()
    try:
        person_birthdate=datetime.date.fromisoformat(birthdate)
    except:
        return True
    if (today-person_birthdate).days > 6574:
        return True
    return False
    
def get_mail(S, c_id, context = "cases"): #Returns mail of person with c_id being case_id or contact_id (number in table, not uuid)
    l = [d for d in S["person_contact_details"].values() if d["person_id"] == S[context][c_id]["person_id"]]
    if len([d for d in l if d["personcontactdetailtype"] == "EMAIL" and d["primarycontact"] == "t"])==1:
        return [d["contactinformation"] for d in l if d["personcontactdetailtype"] == "EMAIL" and d["primarycontact"] == "t"][0]
    elif len([d for d in l if d["personcontactdetailtype"] == "EMAIL"])>0:
        return [d["contactinformation"] for d in l if d["personcontactdetailtype"] == "EMAIL"][0]
    else:
        return "" #Keine Mail im System

def tasks_completed_on_date(S,startdatetime,enddatetime,taskstatus="DONE",tasktype="CASE_INVESTIGATION"): 
    #Output list of task ids with investigation tasks marked as taskstatus on specified date. Date format "YYYY-MM-DD" as String
    #Taskstatus: "DONE", "PENDING", "NOT_EXECUTABLE", "REMOVED"
    completed=[]
    for key in S["tasks"]:
        if (S["tasks"][key]["tasktype"]==tasktype) & (S["tasks"][key]["taskstatus"]==taskstatus) & (S["tasks"][key]["statuschangedate"]>=startdatetime) & (S["tasks"][key]["statuschangedate"]<=enddatetime) :
            completed.append(key)
    return completed

def get_earliest_positive_sample(S,c_id): #Output first positive test for c_id (Output=dict) merged with sample_test entry belonging to sample
    sample_list=[S["samples"][key] for key in S["samples"].keys() if S["samples"][key]["associatedcase_id"]==c_id]
    if len(sample_list)==0:
        return dict([(key,"") for key in next(iter(S["sample_tests"].values())).keys()]) | dict([(key,"") for key in next(iter(S["samples"].values())).keys()])
    else:
        firstsample = min(sample_list,key=lambda d: d["sampledatetime"])
        sampletest_list=[S["sample_tests"][key] for key in S["sample_tests"].keys() if S["sample_tests"][key]["sample_id"]==firstsample["id"]]
        if len(sampletest_list)==0:
            print("Kein Test angelegt")
            return firstsample | dict([(key,"") for key in next(iter(S["samples"].values())).keys()])
        else:
            return firstsample | sampletest_list[0]

def get_earliest_positive_PCR_date(S,c_id): #Output first positive PCR for c_id (Output=dict) merged with sample_test entry belonging to sample
    sample_list=[S["samples"][key] for key in S["samples"].keys() if S["samples"][key]["associatedcase_id"]==c_id]
    if len(sample_list)==0:
        return ""
    else:
        firstsample = min(sample_list,key=lambda d: d["sampledatetime"])
        sampletest_list=[S["sample_tests"][key] for key in S["sample_tests"].keys() if S["sample_tests"][key]["sample_id"]==firstsample["id"] and S["sample_tests"][key]["testtype"]=="PCR_RT_PCR"]
        if len(sampletest_list)==0:
            print("❗ Kein PCR vorhanden")
            return ""
        else:
            return firstsample.get("sampledatetime")

if __name__ == "__main__":
    input_dir = "sormas_export"
    S = read_folder(input_dir)
    print("Alle Dateien eingelesen")
    #S["cases"]: Enthält alle 
    