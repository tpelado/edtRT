from __future__ import print_function

import requests as re
import pdfrw as p
import pypdftk
import subprocess
import xml.dom.minidom as x
from json import dumps
import re
import datetime as dt
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json


#TODO LIST
# nettoyer le code
# gérer le cas du sport


def getPDF():
    result = re.get("https://stri-online.net/Gestion_STRI/TAV/M1/EDT_STRI2_M1.pdf")
    with open('edt.pdf','wb') as f:
        f.write(result.content)

def pdfToTXTfile(filepath):
    f = open("out.txt",'w')
    subprocess.run(["pdf2txt.py",filepath],text=True,stdout=f,check=True)
    f.close()

def pdfToXML(filepath):
    f = open("out.xml",'w')
    subprocess.run(["pdf2txt.py",'-t','xml',filepath],text=True,stdout=f,check=True)
    f.close()

t = ""

# pdfToTXTfile("out.pdf")
# pdfToXML("out.pdf")
# f = open("lol.xml",'r')
# tab = f.readlines()

class semaine:
    id = None
    box = None
    def __init__(self,id,box):
        self.id = id
        self.box = box
    def printSemaine(self):
        print(f"semaine {self.id} boite {self.box}")


class boxClass:
    id = None
    box = None
    value = None
    parentBox = ""
    def __init__(self,id,box,value,parentBox):
        self.id = id
        self.box = [float(i) for i in box.split(",")]
        self.value = value
        self.parentBox = [float(i) for i in parentBox.split(",")]
    def printBox(self):
        print(f"boite:\n\t {self.id}\n boite:\n\t {self.box}\n value:\n\t {self.value}\n parent:\n\t {self.parentBox}")


class coursClass:
    jour = ""
    heureDepart = ""
    heureFin = ""
    duree = ""
    matiere = ""
    salle = ""
    prof = ""
    def __init__(self, jour, heureDepart, heureFin, matiere, salle, prof):
        self.jour = jour
        self.heureDepart = heureDepart
        self.heureFin = heureFin
        self.matiere = matiere
        self.salle = salle
        self.prof = prof

    def printCours(self):
        print(f"Cours de {self.matiere} le {self.jour} de {self.heureDepart} à {self.heureFin} en {self.salle} par {self.prof}")

def createEventObject(summary, location, description, start, end):
    event = {
    'summary': summary,
    'location': location,
    'description': description,
    'start': {
        'dateTime': start,
        'timeZone': 'Europe/Paris',
    },
    'end': {
        'dateTime': end,
        'timeZone': 'Europe/Paris',
    },
    'reminders': {
        'useDefault': False,
    },
    }

    return event

#id du BON CALENDAR
idC = "7auhv9oniguke3igbpmuuqlv9c@group.calendar.google.com"

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def addToGcal(evt):
    try:
        """Shows basic usage of the Google Calendar API.
        Prints the start and name of the next 10 events on the user's calendar.
        """
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        service.events().insert(calendarId=idC,body=evt).execute()
    except ImportError:
        pass

def removeAllFromCal():
    try:
        """Shows basic usage of the Google Calendar API.
        Prints the start and name of the next 10 events on the user's calendar.
        """
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('calendar', 'v3', credentials=creds)
        #TODO FIX
        # Call the Calendar API


        page_token = None
        events = service.events().list(calendarId=idC, pageToken=page_token).execute()
        print(events)
        page_token = events.get('nextPageToken')

        while page_token is not None:
            print(events)
            for event in events['items']:
                service.events().delete(calendarId=idC, eventId=event["id"]).execute()
            events = service.events().list(calendarId=idC, pageToken=page_token).execute()
            page_token = events.get('nextPageToken')




    except ImportError:
        pass




tab = []

doc = x.parse("out.xml")
pages = doc.getElementsByTagName("page")
for page in pages:
    boxes = page.getElementsByTagName("textbox")
    for box in boxes:
        textlines = box.getElementsByTagName("textline")
        for textline in textlines:
            texts = textline.getElementsByTagName("text")
            for text in texts:
                t += text.firstChild.data
            tab.append(boxClass(box.getAttribute("id"),textline.getAttribute("bbox"),t[:-1],box.getAttribute("bbox")))
            t = ""

#<textline bbox="108.000,500.620,129.336,510.238">
#bbox : up left, low left, up right, low right


#box : a boxClass object
#heure : a list of boxes corresponding to hours with 8h first and 19h last
def getHeureCours(box, heure):  # checks for the alignement with known hours "lines" to guess at what time the course starts and ends
    heureDepart = ""
    heureFin  = ""
    x1,x2,x3,x4 = box.parentBox
    for h in heure:
        h1,h2,h3,h4 = h.box
        if x1 <= h1 and heureDepart == "":
            heureDepart = h.value
        if x3 <= h1 and heureFin == "":
            heureFin = h.value



    if heureDepart == "8h":
        heureDepart = "07h45"
    if heureFin == "9h":
        heureFin = "09h45"
    if heureFin == "11h":
        heureFin = "12h05"
    if heureDepart == "10h":
        heureDepart = "10h05"
    if heureDepart == "14h":
        heureDepart = "13h30"
    if heureDepart == "15h":
        heureDepart = "15h45"
    if heureDepart == "16h":
        heureDepart = "15h45"
    if heureDepart == "17h":
        heureDepart = "16h30"
    if heureFin == "15h":
        heureFin = "15h30"
    if heureFin == "17h":
        heureFin = "17h45"
    if heureFin == "18h":
        heureFin = "18h30"

    if heureFin == "10h":
        heureFin = "09h45"
    if heureFin == "12h":
        heureFin = "12h05"
    if heureFin == "16h":
        heureFin = "16h00"



    return heureDepart, heureFin

def getNbrJour(pr):
    dico = {}
    mois = pr[3:]
    dico["janv"] = 31
    dico["feb"] = (29,28)[dt.date.year == 2020]
    dico["mar"] = 31
    dico["avr"] = 30
    dico["mai"] = 31
    dico["juin"] = 30
    dico["juil"] = 31
    dico["nov"] = 30
    dico["déc"] = 31
    return dico[mois]

def getMonthNbr(pr):
    dico = {}
    mois = pr
    dico["janv"] = 1
    dico["feb"] = 2
    dico["mar"] = 3
    dico["avr"] = 4
    dico["mai"] = 5
    dico["juin"] = 6
    dico["juil"] = 7
    dico["aout"] = 8
    dico["sep"] = 9
    dico["oct"] = 10
    dico["nov"] = 11
    dico["déc"] = 12
    return dico[mois]

def getNextMonth(pr):
    mois = pr[3:]
    l = ["janv", "feb", "mar", "avr","mai","juin","juil","aout","sep","oct","nov","déc"]
    i = 0
    for m in l:
        if mois == l:
            return l[(i+1%11)]

        i += 1


def getYear(pr): # TODO fix quand j'aurais la foi
    if(pr == 'nov' or pr == "déc"):
        return 2019
    else:
        return 2020




def getDate(pr, jour):
    jour = jour.value
    mois = pr[3:]
    date = int(pr[0:2])
    # print(jour, date)
    if jour == "Lundi":
        date = date
    if jour == "Mardi":
        date = date+1
        if date > getNbrJour(pr):
            date = 1
            mois = getNextMonth(pr)
    if jour == "Mercredi":
        date = date+2
        if date > getNbrJour(pr):
            date = 2
            mois = getNextMonth(pr)
    if jour == "Jeudi":
        date = date+3
        if date > getNbrJour(pr):
            date = 3
            mois = getNextMonth(pr)
    if jour == "Vendredi":
        date = date+4
        if date > getNbrJour(pr):
            date = 4
            mois = getNextMonth(pr)
    # print(f"{date} {mois}")
    return f"{jour} {date} {mois}"




def getJourCours(box, prJourSemaine, mois):
    i = 0
    jourListe = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi"]
    regex = re.compile("[0-9][0-9]\/[a-zé]{3}")
    date = re.compile("\([0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}\)")
    salle = re.compile("[UK][0-9]{0,2}\-[Aa0-9]{1,3}")

    if (len(box.value) <= 3) or regex.match(box.value) or (box.value in jourListe) or ("EMPLOI DU TEMPS" in box.value) or date.match(box.value) or salle.match(box.value) or "Durant" in box.value or "présence" in box.value or "Pour" in box.value:
        return (-1,None)
    for semaine in mois:
        x1,x2,x3,x4 = box.box
        for jour in semaine.values():
            # jour.printBox()
            j1,j2,j3,j4 = jour.box


            if j4 > x2  and j4 < x4: # cours G1
                return (1,getDate(prJourSemaine[i],jour)) # prof et salle même ligne
            elif j2 > x2 and j2 < x4: #cours G2
                return (2,getDate(prJourSemaine[i],jour)) # idem mais 2nd ligne
            elif j2 > x2 and j4 < x4: # cours g1+g2
                return (3,getDate(prJourSemaine[i],jour)) # prof et salle 2e ligne
        i += 1


def getJourBox(box, prJourSemaine, mois):
    i = 0



    for semaine in mois:

        x1,x2,x3,x4 = box.box
        for jour in semaine.values():
            # jour.printBox()
            j1,j2,j3,j4 = jour.box

            if j4 > x2  and j4 < x4: # cours G1
                return getDate(prJourSemaine[i],jour) # prof et salle même ligne
            elif j2 > x2 and j2 < x4: #cours G2
                return getDate(prJourSemaine[i],jour) # idem mais 2nd ligne
            elif j2 > x2 and j4 < x4: # cours g1+g2
                return getDate(prJourSemaine[i],jour) # prof et salle 2e ligne
        i += 1

#retourne le jour est l'heure d'une case de
def getProfBox(box, prJourSemaine, mois):
    i = 0

    date = re.compile("\([0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}\)")
    salle = re.compile("[UK][0-9]{0,2}\-[Aa0-9]{1,3}")
    prof = re.compile("[A-Z\/]+")

    if date.match(box.value) or salle.match(box.value):
        return -1

    if prof.match(box.value):
        for semaine in mois:
            x1,x2,x3,x4 = box.box
            for jour in semaine.values():
                # jour.printBox()
                j1,j2,j3,j4 = jour.box

                if j4 > x2  and j4 < x4: # cours G1
                    return 1,getDate(prJourSemaine[i],jour) # prof et salle même ligne
                elif j2 > x2 and j2 < x4: #cours G2
                    return 2,getDate(prJourSemaine[i],jour) # idem mais 2nd ligne
                elif j2 > x2 and j4 < x4: # cours g1+g2
                    return 3,getDate(prJourSemaine[i],jour) # prof et salle 2e ligne
            i += 1



i = 8
heure = {}
#on récupére les coords de 8h, 9h, 10h...for box in tab:
for box in tab:
    if f"{i}h" in box.value and f"{i+10}'" not in box.value:
           # box.printBox()
            heure[f"{i}h"] = box
            i += 1

#on va récup la position de chaque jour de chaque semaine...
mois = [] # liste de semaine
prJourSemaine = [] # liste des premiers jours de la semaine
semaine = {} #dico de jour

jourListe = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi"]

for box in tab:
    for j in jourListe:
        if j in box.value:
            semaine[j] = box
            if j == "Vendredi":
                mois.append(semaine)
                semaine = {}


regex = re.compile("[0-9][0-9]\/[a-zé]{3}")

for box in tab:
    if regex.match(box.value):
        prJourSemaine.append(box.value)


salle = re.compile("[UK][0-9]{0,2}\-[Aa0-9]{1,3}")

tabSalles = []
for box in tab:
    if salle.match(box.value):
        tabSalles.append(box)


tabProf = []
prof = re.compile("[A-Z\/]+")
for box in tab:
    if prof.fullmatch(box.value):
        tabProf.append(box)



listeCoursSuivi = ["BD/WD"]
edt = []
salle = None
prof = None
for box in tab:
        h1, h2 = getHeureCours(box, heure.values())
        try:
            (groupe,jour) = getJourCours(box, prJourSemaine, mois)
        except TypeError:
            box.printBox()
            break
        if groupe == 1:
            # la salle et le prof sont sur la même ligne 1
            for s in tabSalles:
                sj = getJourBox(s, prJourSemaine, mois)
                hs1, hs2 = getHeureCours(s, heure.values())
                # print(h2, hs2)
                if sj == jour and hs2 == h2:
                    # print("===========",jour,sj)
                    # s.printBox()
                    salle = s
                    break
            for p in tabProf:
                (gr,pj) = getProfBox(p, prJourSemaine, mois)

                if pj == jour and gr == 2 and " - " not in box.value:
                    prof = p
                    break
                else:
                    prof = None
            if prof is not None:
                edt.append(coursClass(jour,h1,h2 ,box.value,salle.value , prof.value ))
            else:
                prof = box.value[box.value.find("-")+2:]
                edt.append(coursClass(jour,h1,h2 ,box.value[:box.value.find("-")-1],salle.value ,prof ))







print("Agenda Groupe 1 non ordonné")

removeAllFromCal()


# for c in edt:
#     c.printCours()
#     d = c.jour[c.jour.rfind(' ')+1:]
#     j = c.jour[c.jour.find(' '):][1:]
#     j = j[:j.find(' ')]
#     dateStr=f"{getYear(d)}-{getMonthNbr(d)}-{j} {c.heureDepart.replace('h',':')}"
#     print(dateStr)
#     dateDebut = datetime.datetime.strptime(dateStr, '%Y-%m-%d %H:%M')
#     dateStr=f"{getYear(d)}-{getMonthNbr(d)}-{j} {c.heureFin.replace('h',':')}"
#     print(dateStr)
#     dateFin = datetime.datetime.strptime(dateStr, '%Y-%m-%d %H:%M')
#     event = createEventObject(c.matiere,c.salle,"cours",dateDebut.isoformat(),dateFin.isoformat())
#     print(event)
#     addToGcal(event)
#


















