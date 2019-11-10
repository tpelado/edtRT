import xml.dom.minidom as x



class boxClass:
    id = None
    box = None
    value = None
    parentBox = ""
    def __init__(self,id,box,value,parentBox):
        self.id = id
        self.box = [float(i) for i in box.split(",")]
        self.value = value
        self.parentBox = parentBox
    def printBox(self):
        print(f"boite:\n\t {self.id}\n boite:\n\t {self.box}\n value:\n\t {self.value}\n parent:\n\t {self.parentBox}")



def parseTextLine():
    t = ""
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
    doc = x.parse("out.xml")
    pages = doc.getElementsByTagName("page")
    for page in pages:
        layout = page.getElementsByTagName("layout")
        elt = layout.item(0)
        parseTextgroup(elt.childNodes[1], None)
    return tab

tabLayout = []
def parseTextgroup(elt, parentElt):
    global tabLayout
    id = []
    e = elt.getElementsByTagName("textgroup")
    if e == []:
        for tb in elt.getElementsByTagName("textbox"):
            id.append(int(tb.getAttribute("id")))
        tabLayout.append(boxClass(id, elt.getAttribute("bbox"),"value",parentElt))
        id = ""
    else:
        for tb in elt.getElementsByTagName("textbox"):
             id.append(int(tb.getAttribute("id")))
        tabLayout.append(boxClass(id, elt.getAttribute("bbox"),"value",parentElt))
        id = []
        for el in e:
            parseTextgroup(el, e)



def isIn(sl, l):
    return set(sl).issubset(l)



























