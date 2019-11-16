
from __future__ import print_function

import requests as req
import subprocess
import xml.dom.minidom as x
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
#osekour


###### Ce code est moche et pas commenté. Il est aussi relativement imbitable. 

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
    matiere = ""
    salle = ""
    prof = ""
    idGrp = []
    page = None
    box = None
    
    def __init__(self, jour, heureDepart, heureFin, matiere, salle, prof, idGrp, page, boxe):
        self.jour = jour
        self.heureDepart = heureDepart
        self.heureFin = heureFin
        self.matiere = matiere
        self.salle = salle
        self.prof = prof
        self.idGrp = idGrp
        self.page = page
        
        self.box = boxe
        if self.box is None:
            exit()
   

    
    
    def printCours(self):
        print(f"Cours de {self.matiere} le {self.jour} de {self.heureDepart} à {self.heureFin} en {self.salle} par {self.prof}\n idGrp = {self.idGrp}\n page = {self.page}\n")
        print("=========\n box =" )
        self.box.printBox()
        print("\n =====FIN======")






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

        service = build('calendar', 'v3', credentials=creds, cache_discovery=False)

        # Call the Calendar API
        service.events().insert(calendarId=idC,body=evt).execute()
    except (ImportError,ModuleNotFoundError):
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

        service = build('calendar', 'v3', credentials=creds, cache_discovery=False)
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
    except (ImportError,ModuleNotFoundError):
        pass


#### Program functions


def getPDF(): # get a updated version of the online PDF
    result = req.get("https://stri-online.net/Gestion_STRI/TAV/M1/EDT_STRI2_M1.pdf")
    with open('edt.pdf','wb') as f:
        f.write(result.content)


def pdfToXML(filepath):#convert a pdf file to a readable xml file via the pdf2txt.py python app
    f = open("out.xml",'w')
    subprocess.run(["pdf2txt.py",'-t','xml',filepath],text=True,stdout=f,check=True)
    f.close()

def updatePDF():
    getPDF()
    pdfToXML("edt.pdf")
    



#### parse le document xml


def parseTextLine(): # main parsing function to get something out of the XML file. Can't be bothered to explain this one, its quite simple if you look at the xml
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


def parseTextgroup(elt, parentElt, page): # recursive private function for parseTextLine
    global tabLayout
    id = []
    e = elt.getElementsByTagName("textgroup")
    if e == []:
        for tb in elt.getElementsByTagName("textbox"):
            id.append(int(tb.getAttribute("id")))
        #we don't need the parent box and the value so those are dummies
        tabLayout[page].append(boxClass(id, elt.getAttribute("bbox"),"value","0.0 ,0.0 , 0.0, 0.0", page))
        id = ""
    else:
        for tb in elt.getElementsByTagName("textbox"):
             id.append(int(tb.getAttribute("id")))
        
        tabLayout[page].append(boxClass(id, elt.getAttribute("bbox"),"value","0.0 ,0.0 , 0.0, 0.0", page))
        id = []
        for el in e:
            parseTextgroup(el, e,page)



def upscaleHour(heure, res): # takes a list of boxes representing hours (8h, 9h...) and adds inbetween hours (8h10, 8h20...) 
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
            nheure[f"{heure[l[i]].value.zfill(3)}{str(j).zfill(2)}"] = boxClass(None,f"{pos + (BtoA*(k/res))},{heure[l[i]].box[1]},{pos + (BtoA*(k/res))+size},{heure[l[i]].box[3]}", None, "0.0,0.0,0.0,0.0", None)
            j += int(60/res)
            k += 1
        j = 0
        k = 0

    nheure["19h00"] = boxClass(None,f"794.0625, 510.13, 805.8125, 520.225", None, "0.0,0.0,0.0,0.0",None)
    
    nheure2 = {}
    
    for k in sorted(nheure.keys()):
        nheure2[k] = nheure[k]
        
        
    return nheure2





#box : a boxClass object
#heure : a dict of boxes corresponding to hours with 8h first and 19h last
def getHeureCours(box, heure):  # checks for the alignement with known hours "lines" to guess at what time the course starts and ends
    heureDepart = ""
    heureFin  = ""
    x1,x2,x3,x4 = box.box
    for h in heure:
        h1,h2,h3,h4 = heure[h].box
        # print(h)
        # print(x1,"<=", h1, x3,"<=", h1)
        if x1 <= h1 and heureDepart == "":
            heureDepart = h
        if x3 <= h1 and heureFin == "":
            heureFin = h
    return heureDepart, heureFin

#used only once. This is horrible but if I remove it it WILL break everything. 
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




def getDate(pr, jour): # get the date for a day of the week and the first day of the week. could be done with a list but it works
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



def strInList(str, list): # checks whether a string is part of a string in a list of strings
    for e in list:
        if e in str:
            return True
    return False



def getJourCours(box, prJourSemaine, mois):#get the boxes containing the days of the week to check for alignent and know what course is when
    i = 0
    jourListe = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi"]
    regex = re.compile("[0-9][0-9]\/[a-zé]{3}")
    date = re.compile("\([0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}\)")
    salle = re.compile("[UK][0-9]{0,2}\-[Aa0-9]{1,3}")
    listeNoire = ["Durant", "présence", "Pour", "EMPLOI"]
    if (len(box.value) <= 3) or regex.match(box.value) or (box.value in jourListe) or date.match(box.value) or salle.match(box.value) or strInList(box.value, listeNoire):
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





def ajouterLesCours(mois, prJourSemaine, heure,nheure, tab): # create a course from the boxes and adds a prof and room to it if it finds one
    edt = []
    salle = None
    prof = None
    
    for box in tab:
            h1, h2 = getOldHeureCours(box, heure.values()) # get an idea of when the course starts and ends
            try:
                (groupe,jour) = getJourCours(box, prJourSemaine, mois) # gets a day for the course
                if groupe == 1: # if the course is for us in group 1
                
                    aligned, salle = determineClosestFromBox(box, tabSalles) # we check if there is a room aligned with our course, if yes it means the prof name is in the course name. if not it means the prof name is under the course name
                    if not aligned:
                        osef, prof = determineClosestFromBox(box, tabProf) # we try to check whether or not we can get the prof name
                        if salle is not None and prof is not None:
                            if abs(prof.box[1] - salle.box[1]) > 10: # if the prof and room are not aligned, we're using the wrong room or room, so we discard them
                                prof = None
                                salle = None
                                
                    # we add the course to the timetable
                    
                    # we check if there is a "-" in the course name (meaning the prof value is wrong since there's no prof associated to this course)
                    # we don't want to check for the "sport" course since it has no prof or room and we don't really care about this course anyway
                    if salle is not None and prof is not None and "-" not in box.value and "Sport" not in box.value:
                        edt.append(coursClass(jour,h1,h2 ,box.value,salle.value , prof.value,[box.id, prof.id],box.page, box ))
                    
                    # we search the prof name in the course name and add the course 
                    elif "-" in box.value and salle is not None:
                        prof = box.value[box.value.find("-")+2:]
                        edt.append(coursClass(jour,h1,h2 ,box.value[:box.value.find("-")-1],salle.value ,prof ,[box.id],box.page, box ))
                        prof = None
                    
                    
                    elif salle is not None:
                        try:
                            edt.append(coursClass(jour,h1,h2 ,box.value,salle.value ,prof ,[box.id, prof.id],box.page, box ))
                        except AttributeError:
                            edt.append(coursClass(jour,h1,h2 ,box.value,salle.value ,None,[box.id],box.page,box ))
                
                    else:
                        edt.append(coursClass(jour,h1,h2 ,box.value,None,None,[box.id],box.page,box ))


            except TypeError:
                print("THIS BOX HAS CAUSED AN EXECPTION PLEASE HELP ME GOD")
                box.printBox()
                
            
    return edt



def isIn(sl, l): # checks if a sublist is part of a list
    return set(sl).issubset(l)


def deDup(liste): # this will remove any duplicates from a list
    return list(set(liste))


def computeDistance(box1, box2): # compute the x and y distance between the middle of box 2 and the middle of box 1
    x1, x2, x3, x4 = box1.box
    w1, w2, w3, w4 = box2.box
    meanBox1 = ((x1+x3)/2,(x2+x4)/2)
    meanBox2 = ((w1+w3)/2,(w2+w4)/2)
    x,y = abs(meanBox1[0] - meanBox2[0]),abs(meanBox1[1] - meanBox2[1])
    # print(f"Mean of box 1 is {meanBox1}, Mean of box2 is {meanBox2}")
    # print(f"absolute distance is {x} in x and {y} in y")
    return x,y



def isClosest(xy1, xy2): # checks if xy1 tuple is closer than xy2 tuple.
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










def getTextGroup(tabLayout, idList, page): # this takes the smallest box with all boxes id (generally the prof and box id) to have the best possible box to guess at what time the course starts and ends
    idList = [int(i) for i in deDup(idList)]
    minBox = -1
    minBoite = None
    # print(idList)
    for box in tabLayout[page]:
        if isIn(idList, box.id):
            if minBox == -1 or len(box.id) < minBox :
                minBox = len(box.id)
                minBoite = box
    return minBoite




dryRun = False # used to see courses withouth adding them to the calendar

###################### main program

if not dryRun:
    updatePDF()



tabBox = parseTextLine()
mois, prJourSemaine, tabSalles, tabProf, heure = getPosDays(tabBox)
nheure = upscaleHour(heure, 12)
edt = ajouterLesCours(mois, prJourSemaine,heure, nheure, tabBox)





if not dryRun:
    removeAllFromCal()





edt = list(set(edt))

if True:
    matSuivie=["Réseaux Mobiles","Sécurité","Java","Gestion Réseaux","Anglais","BD/WD", "Interco"]
    # matSuivie=["Java"]
    
    
    for c in edt:
        if strInList(c.matiere,matSuivie):
            
            boiteCours = getTextGroup(tabLayout, c.idGrp, c.page)
            c.heureDepart, c.heureFin = getHeureCours(boiteCours, nheure)
            
            if c.heureDepart == "10h00" and c.heureFin == "14h30": # fix pour un autre cours pété dans un groupe de boite encore plus pété 
                c.heureDepart = "13h30"
            
            hd = int(c.heureDepart[:2])
            hf = int(c.heureFin[:2])
            
            if hf-hd > 3 or hf-hd < 2:
                c.heureFin = str(int(c.heureDepart[:2]) + 2)+c.heureDepart[2:]
        
                
            
            if c.heureDepart == "12h00":
                c.heureDepart = "13h30" # fix pour un cours pété dans un groupe de boite pété
            
          
           
            if c.heureDepart == "19h00":
                pass # help 
                
            
                # edt.remove(c)
            else:
                # print("==============")
                # c.printCours()
                d = c.jour[c.jour.rfind(' ')+1:]
                j = c.jour[c.jour.find(' '):][1:]
                j = j[:j.find(' ')]
                dateStr=f"{getYear(d)}-{getMonthNbr(d)}-{j} {c.heureDepart.replace('h',':')}"
                # print(dateStr)
                dateDebut = datetime.datetime.strptime(dateStr, '%Y-%m-%d %H:%M')
                dateStr=f"{getYear(d)}-{getMonthNbr(d)}-{j} {c.heureFin.replace('h',':')}"
                # print(dateStr)
                dateFin = datetime.datetime.strptime(dateStr, '%Y-%m-%d %H:%M')
                try:
                    event = createEventObject(c.matiere,c.salle,c.salle,dateDebut.isoformat(),dateFin.isoformat())
                except ValueError:
                    pass
                # print(event)
                if not dryRun:
                    addToGcal(event)



















