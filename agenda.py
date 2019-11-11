from __future__ import print_function

import requests as re
import pdfrw as p
import pypdftk
import subprocess
import xml.dom.minidom as x
from json import dumps
import re
import datetime as dt
from time import sleep
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json
import sys
from math import sqrt

#TODO LIST
# nettoyer le code
# gérer le cas du sport




####global variables
tabLayout = [[]]

####classes
class boxClass: # a box class generated from a textbox or textgroup xml object
    id = None
    box = None
    value = None
    page = None
    parentBox = ""
    def __init__(self,id,box,value,parentBox, page):
        self.id = id
        self.box = [float(i) for i in box.split(",")]
        self.value = value
        self.parentBox = [float(i) for i in parentBox.split(",")]
        self.page = page
    def printBox(self):
        print(f"boite:\n\t {self.id}\n boite:\n\t {self.box}\n value:\n\t {self.value}\n parent:\n\t {self.parentBox} \n page {self.page}")


class coursClass: # a course object, with all the info we can get to setup the course on gcalendar
    jour = ""
    heureDepart = ""
    heureFin = ""
    duree = ""
    matiere = ""
    salle = ""
    prof = ""
    idGrp = []
    page = None
    def __init__(self, jour, heureDepart, heureFin, matiere, salle, prof, idGrp, page):
        self.jour = jour
        self.heureDepart = heureDepart
        self.heureFin = heureFin
        self.matiere = matiere
        self.salle = salle
        self.prof = prof
        self.idGrp = idGrp
        self.page = page
    def printCours(self):
        print(f"Cours de {self.matiere} le {self.jour} de {self.heureDepart} à {self.heureFin} en {self.salle} par {self.prof}\n idGrp = {self.idGrp}\n page = {self.page}")






#### google calendar interaction

def createEventObject(summary, location, description, start, end): #create an event object to add in gcalendar
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

#id for the calendar to modify
idC = "7auhv9oniguke3igbpmuuqlv9c@group.calendar.google.com"

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']



def addToGcal(evt):#add an event
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

def removeAllFromCal(): # clear the calendar
    try:
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
        # print(events)
        page_token = events.get('nextPageToken')

        while page_token is not None:
            # print(events)
            for event in events['items']:
                service.events().delete(calendarId=idC, eventId=event["id"]).execute()
            page_token = events.get('nextPageToken')
            events = service.events().list(calendarId=idC, pageToken=page_token).execute()
    except ImportError:
        pass


#### Program functions


def getPDF(): # get a updated version of the online PDF
    result = re.get("https://stri-online.net/Gestion_STRI/TAV/M1/EDT_STRI2_M1.pdf")
    with open('edt.pdf','wb') as f:
        f.write(result.content)


def pdfToTXTfile(filepath):#deprecated, use pdfToXML
    f = open("out.txt",'w')
    subprocess.run(["pdf2txt.py",filepath],text=True,stdout=f,check=True)
    f.close()


def pdfToXML(filepath):#convert a pdf file to a readable xml file via the pdf2txt.py python app
    f = open("out.xml",'w')
    subprocess.run(["pdf2txt.py",'-t','xml',filepath],text=True,stdout=f,check=True)
    f.close()




#### parse le document xml


def parseTextLine():
    global tabLayout
    t = ""
    i = 0
    tab = []
    pos = None
    flag = True
    doc = x.parse("out.xml")
    pages = doc.getElementsByTagName("page")
    for page in pages:
        tabLayout.append([])
        boxes = page.getElementsByTagName("textbox")
        for box in boxes:
            textlines = box.getElementsByTagName("textline")
            for textline in textlines:
                texts = textline.getElementsByTagName("text")
                for text in texts:
                    if flag:
                        flag = False
                        pos = text.getAttribute("bbox")
                    t += text.firstChild.data
                # tab.append(boxClass(box.getAttribute("id"),textline.getAttribute("bbox"),t[:-1],,i))
                tab.append(boxClass(box.getAttribute("id"),pos,t[:-1],box.getAttribute("bbox"),int(page.getAttribute("id"))))
                flag = True
                pos = None
                t = ""
        i += 1



    pages = doc.getElementsByTagName("page")
    i = 0
    for page in pages:
        layout = page.getElementsByTagName("layout")
        elt = layout.item(0)
        parseTextgroup(elt.childNodes[1],"0.0 ,0.0 , 0.0, 0.0",int(page.getAttribute("id")))
        i += 1
    return tab

#TODO fix parent box
def parseTextgroup(elt, parentElt, page): # recursive private function for parseTextLine
    global tabLayout
    id = []
    e = elt.getElementsByTagName("textgroup")
    if e == []:
        for tb in elt.getElementsByTagName("textbox"):
            id.append(int(tb.getAttribute("id")))
        # tabLayout.append(boxClass(id, elt.getAttribute("bbox"),"value",parentElt))
        tabLayout[page].append(boxClass(id, elt.getAttribute("bbox"),"value","0.0 ,0.0 , 0.0, 0.0", page))
        id = ""
    else:
        for tb in elt.getElementsByTagName("textbox"):
             id.append(int(tb.getAttribute("id")))
        # tabLayout.append(boxClass(id, elt.getAttribute("bbox"),"value",parentElt))
        tabLayout[page].append(boxClass(id, elt.getAttribute("bbox"),"value","0.0 ,0.0 , 0.0, 0.0", page))
        id = []
        for el in e:
            parseTextgroup(el, e,page)



def upscaleHour(heure, res):
    #heure is dict

    j = 0
    k = 0
    nheure = {}
    nheure["07h45"] = boxClass(None,f"108.0,510.13,113.0,520.225", None, "0.0,0.0,0.0,0.0",None)
    l = list(heure.keys())
    BtoA = 0
    for i in range(0, len(l)-1):
        pos, pos2 = heure[l[i]].box[0], heure[l[i+1]].box[0]
        size = heure[l[i]].box[2] - heure[l[i]].box[0]
        BtoA = pos2 - pos
        while j < 60:
            # print(f"{heure[l[i]].value.zfill(2)}{str(j).zfill(2)}", f"{pos + (BtoA*(k/res))},{heure[l[i]].box[1]},{pos + (BtoA*(k/res))+size},{heure[l[i]].box[3]}")
            nheure[f"{heure[l[i]].value.zfill(2)}{str(j).zfill(2)}"] = boxClass(None,f"{pos + (BtoA*(k/res))},{heure[l[i]].box[1]},{pos + (BtoA*(k/res))+size},{heure[l[i]].box[3]}", None, "0.0,0.0,0.0,0.0", None)
            j += int(60/res)
            k += 1
        j = 0
        k = 0

    nheure["19h00"] = boxClass(None,f"794.0625, 510.13, 805.8125, 520.225", None, "0.0,0.0,0.0,0.0",None)
    return nheure





#box : a boxClass object
#heure : a dict of boxes corresponding to hours with 8h first and 19h last
def getHeureCours(box, heure):  # checks for the alignement with known hours "lines" to guess at what time the course starts and ends
    heureDepart = ""
    heureFin  = ""
    x1,x2,x3,x4 = box.box
    for h in heure:
        h1,h2,h3,h4 = heure[h].box
        if x1 <= h1 and heureDepart == "":
            heureDepart = h
        if x3 <= h1 and heureFin == "":
            heureFin = h
    return heureDepart, heureFin


def getOldHeureCours(box, heure):  # checks for the alignement with known hours "lines" to guess at what time the course starts and ends
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





def getNbrJour(pr): # get the number of day in a month based on the month name
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

def getMonthNbr(pr):# get the month number based on the month name
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

def getNextMonth(pr): # get the next month after the given one
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




def getDate(pr, jour): # get the date for a day of the week and the first day of the week
    jour = jour.value
    mois = pr[3:]
    date = int(pr[0:2])
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
    return f"{jour} {date} {mois}"




def getJourCours(box, prJourSemaine, mois):#get the boxes containing the days of the week to check for alignent and know what course is when
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
            if jour.page == box.page:
                # jour.printBox()
                j1,j2,j3,j4 = jour.box
    
    
                if j4 > x2  and j4 < x4: # cours G1
                    return (1,getDate(prJourSemaine[i],jour)) # prof et salle même ligne
                elif j2 > x2 and j2 < x4: #cours G2
                    return (2,getDate(prJourSemaine[i],jour)) # idem mais 2nd ligne
                elif j2 > x2 and j4 < x4: # cours g1+g2
                    return (3,getDate(prJourSemaine[i],jour)) # prof et salle 2e ligne
        i += 1




def getJourBox(box, prJourSemaine, mois): # ?
    i = 0



    for semaine in mois:

        x1,x2,x3,x4 = box.box
        for jour in semaine.values():
            print("==============================",semaine.page)
            if semaine.page == box.page:
            # jour.printBox()
                j1,j2,j3,j4 = jour.box
                
                if j4 >= x2  and j4 < x4: # cours G1
                    return getDate(prJourSemaine[i],jour) # prof et salle même ligne
                elif j2 > x2 and j2 < x4: #cours G2
                    return getDate(prJourSemaine[i],jour) # idem mais 2nd ligne
                elif j2 > x2 and j4 < x4: # cours g1+g2
                    return getDate(prJourSemaine[i],jour) # prof et salle 2e ligne
            i += 1

def getProfBox(box, prJourSemaine, mois): # return the "day and hour" coords to check what prof. teaches what course and when.
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


                if j4 >= x2  and j4 < x4: # cours G1
                    return 1,getDate(prJourSemaine[i],jour) # prof et salle même ligne
                elif j2 > x2 and j4 > x4 and j2 < x4: #cours G2
                    return 2,getDate(prJourSemaine[i],jour) # idem mais 2nd ligne
                elif j2 > x2 and j4 < x4: # cours g1+g2
                    return 3,getDate(prJourSemaine[i],jour) # prof et salle 2e ligne
            i += 1




def getPosDays(tab):
    #on va récup la position de chaque jour de chaque semaine...
    mois = [] # liste de semaine
    prJourSemaine = [] # liste des premiers jours de la semaine
    semaine = {} #dico de jour
    tabSalles = []
    tabProf = []
    i = 8
    heure = {}

    jourListe = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi"]


    regex = re.compile("[0-9][0-9]\/[a-zé]{3}")
    salle = re.compile("[UK][0-9]{0,2}\-[Aa0-9]{1,3}")
    salleAnglais = re.compile("[K][0-9]{0,2}")
    prof = re.compile("[A-Z\/]+")



    for box in tab:
        if f"{i}h" in box.value and f"{i+10}'" not in box.value:
            # box.printBox()
                heure[f"{i}h"] = box
                i += 1
        if prof.fullmatch(box.value):
            tabProf.append(box)
        if salle.match(box.value) or salleAnglais.match(box.value):
            tabSalles.append(box)
        if regex.match(box.value):
            prJourSemaine.append(box.value)
        for j in jourListe:
            if j in box.value:
                semaine[j] = box
                if j == "Vendredi":
                    mois.append(semaine)
                    semaine = {}


    return mois, prJourSemaine, tabSalles, tabProf, heure





def ajouterLesCours(mois, prJourSemaine, heure,nheure, tab):
    edt = []
    salle = None
    prof = None
    for box in tab:
            h1, h2 = getOldHeureCours(box, heure.values())
            try:
                (groupe,jour) = getJourCours(box, prJourSemaine, mois)
            
            except TypeError:
                box.printBox()
                break
            if groupe == 1:
            
                aligned, salle = determineClosestFromBox(box, tabSalles)
                if not aligned:
                    osef, prof = determineClosestFromBox(box, tabProf)
                    if salle is not None and prof is not None:
                        if prof.box[1] != salle.box[1]: # la salle est le prof ne sont pas alignés : ya un soucis
                            prof = None
                            salle = None
                        
                    
                    
                    
                    # edt.append(coursClass(jour,h1,h2 ,box.value,salle.value , prof.value,[box.id, salle.id, prof.id],box.page ))
                
                
                
                
                
                if salle is not None and prof is not None and "-" not in box.value and "Sport" not in box.value:
                    edt.append(coursClass(jour,h1,h2 ,box.value,salle.value , prof.value,[box.id, prof.id],box.page ))
                   
                elif "-" in box.value and salle is not None:
                    prof = box.value[box.value.find("-")+2:]
                    edt.append(coursClass(jour,h1,h2 ,box.value[:box.value.find("-")-1],salle.value ,prof ,[box.id],box.page ))
                    prof = None
                
                elif salle is not None:
                    # edt.append(coursClass(jour,h1,h2 ,None,salle.value ,prof ,[box.id, salle.id],box.page ))
                    try:
                        edt.append(coursClass(jour,h1,h2 ,box.value,salle.value ,prof ,[box.id, prof.id],box.page ))
                    except AttributeError:
                        edt.append(coursClass(jour,h1,h2 ,box.value,salle.value ,None,[box.id],box.page ))
               
                else:
                    edt.append(coursClass(jour,h1,h2 ,box.value,None,None,[box.id],box.page ))


    return edt



def isIn(sl, l):
    return set(sl).issubset(l)


def deDup(liste):
    return list(set(liste))


def computeDistance(box1, box2): # compute the absolute distance between the middle of box 2 and the middle of box 1
    x1, x2, x3, x4 = box1.box
    w1, w2, w3, w4 = box2.box
    meanBox1 = ((x1+x3)/2,(x2+x4)/2)
    meanBox2 = ((w1+w3)/2,(w2+w4)/2)
    x,y = abs(meanBox1[0] - meanBox2[0]),abs(meanBox1[1] - meanBox2[1])
    # print(f"Mean of box 1 is {meanBox1}, Mean of box2 is {meanBox2}")
    # print(f"absolute distance is {x} in x and {y} in y")
    return x,y

def computeDistance2(box1, box2): # NOT USED compute the absolute distance between the middle of box 2 and the middle of box 1
    x1, x2, x3, x4 = box1.box
    w1, w2, w3, w4 = box2.box
    mb1 = ((x1+x3)/2,(x2+x4)/2)
    mb2 = ((w1+w3)/2,(w2+w4)/2)
    return sqrt( ( ( mb2[0]-mb1[0] )**2) + ( ( mb2[1]-mb1[1] )**2) )




def isClosest(xy1, xy2): # compute if xy1 tuple is less than xy2 tuple.
    closestInX = xy1[0] < xy2[0]
    closestInY = xy1[1] < xy2[1]
    return closestInX, closestInY




def determineClosestFromBox(box, tab): # return the element from the tab which is the closest to the box
    cx, cy = (999999, 999999)
    x,y = (99999,99999)
    closestBox = None
    # box.printBox()
    for b in tab:
        # print(b.page, box.page)

        if (b.box[1]  <= box.box[1] or abs(b.box[1] - box.box[1]) < 2) and (b.box[0] >= box.box[0] or abs(b.box[0] - box.box[0]) < 2)  and ((b.box[1] - box.box[1]) < 50 ) and b.page == box.page :
            x,y = computeDistance(box, b)
            # print(f"{x,y}  {cx, cy}")
            if isClosest((x,y),(cx,cy))[1]: # closer in Y, not in X
                if(x < 200): # we don't want a room aligned with the course but on the other side of the time table
                    cx,cy = x,y
                    closestBox = b
            elif (y == cy) and isClosest((x,y),(cx,cy))[0]: # closer in X but not in Y, y levels are the same
                cx,cy = x,y
                closestBox = b

    
    return cy==0, closestBox






def determineClosestFromBox2(box, tab): # return the element from the tab which is the closest to the box
    distance = 999999
    cy = 0
    d = 1
    closestBox = None
    # box.printBox()
    for b in tab:
        # print(b.page, box.page)
        if (b.page == box.page):
            if (b.box[1]  <= box.box[1] or abs(b.box[1] - box.box[1]) < 2) and (b.box[0] >= box.box[0] or abs(b.box[0] - box.box[0]) < 2) :
                x,y = computeDistance(box, b)
                d = computeDistance2(box, b)
                if d < distance: #
                    d = distance
                    cy = y
                    closestBox = b

    # closestBox.printBox()
    return cy==0, closestBox





def getTextGroup(tabLayout, idList, page):
    idList = [int(i) for i in deDup(idList)]
    minBox = -1
    minBoite = None
    print(idList)
    for box in tabLayout[page]:
        if isIn(idList, box.id):
            if minBox == -1 or len(box.id) < minBox :
                minBox = len(box.id)
                minBoite = box
    print("*****************************************")
    print(idList)
    minBoite.printBox()
    print("***************FIN******************")
    return minBoite



tabBox = parseTextLine()
mois, prJourSemaine, tabSalles, tabProf, heure = getPosDays(tabBox)
nheure = upscaleHour(heure, 12)
edt = ajouterLesCours(mois, prJourSemaine,heure, nheure, tabBox)


# tabSalles[0].printBox()
# tabBox[27].printBox()
# computeDistance(tabSalles[0], tabBox[27])



# for b in tabBox:
#     print(b.value)
#     determineClosestFromBox(b, tabSalles)
#     print("-------------")
#     determineClosestFromBox(b, tabProf)
#     print("===========================================")
#     sleep(2)


# print("Agenda Groupe 1 non ordonné")
removeAllFromCal()


# for s in tabSalles:
#
#     s.printBox()
#     print(getOldHeureCours(s, heure.values()))
#     print(getHeureCours(s, heure))
#     print(getJourBox(s,prJourSemaine,mois))
#     sleep(1)
# 

# for h in nheure:
#     print(h)
#     nheure[h].printBox()

# 
edt = list(set(edt))

for c in edt:
    if c.page == 1:
    # c.printCours()
        boiteCours = getTextGroup(tabLayout, c.idGrp, c.page)
        c.heureDepart, c.heureFin = getHeureCours(boiteCours, nheure)
        c.heureFin = str(int(c.heureDepart[:2]) + 2)+c.heureDepart[2:]
        if c.heureDepart == "19h00":
            pass
            # edt.remove(c)
        else:
            print("==============")
            c.printCours()
            d = c.jour[c.jour.rfind(' ')+1:]
            j = c.jour[c.jour.find(' '):][1:]
            j = j[:j.find(' ')]
            dateStr=f"{getYear(d)}-{getMonthNbr(d)}-{j} {c.heureDepart.replace('h',':')}"
            print(dateStr)
            dateDebut = datetime.datetime.strptime(dateStr, '%Y-%m-%d %H:%M')
            dateStr=f"{getYear(d)}-{getMonthNbr(d)}-{j} {c.heureFin.replace('h',':')}"
            print(dateStr)
            dateFin = datetime.datetime.strptime(dateStr, '%Y-%m-%d %H:%M')
            try:
                event = createEventObject(c.matiere,c.salle,"cours",dateDebut.isoformat(),dateFin.isoformat())
            except ValueError:
                pass
            print(event)
            addToGcal(event)



















