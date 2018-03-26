#cMaker - Calculation Maker
#by Regan Lu

#Libaraies
import numpy as np
import openpyxl as xl
from  Tkinter import *
import Tkinter, Tkconstants, tkFileDialog
import matplotlib.pyplot as mlib

#Variables
loaded = False
exiter = False
starter = False
importv = False
root = Tk()
gFile = 0

#Calculation Class
class var: #variable class
    name = ""   #name
    var = ""    #variable
    value = ""  #value
    unit = ""   #unit
    note = ""   #notes
    ref = ""    #reference
    cell = ""   #location in sheet

    def mCont(self,List,setting):

        tmpList = List
        
        #Latex expression as it is not generated yet
        tmpList = addList(List,"")

        #Cell Reference
        tmpList = addList(List,"")

        return tmpList
    
    def intCont(self,ListV):
        self.name = ListV[0]
        self.var = ListV[1]
        self.value = ListV[2]
        self.unit = ListV[3]
        self.note = ListV[4]
        self.ref = ListV[5]

        #Missing info
        self.cell = ListV[6]
    
class seq: #sequence class
    seq = 0     #sequence order
    typ = 0     #type: 0 - formula, 1 - check/if statement, 2 - HLookup, 3 - Vlookup
    name = ""   #name
    var = ""    #variable
    unit = ""   #unit
    exp = ""    #expression
    lexp = ""   #latex expression
    A = ""      #Control variable A
    B = ""      #Control variable B
    C = ""      #Control variable C
    ref = ""    #Reference
    cell = ""    #location in sheet

    def mCont(self,List,setting):

        tmpList = List
        
        #Latex expression as it is not generated yet
        tmpList = addList(List,"")

        #Cell Reference
        tmpList = addList(List,"")

        return tmpList

    def intCont(self,listS):
        self.seq = listS[0]     #sequence order
        self.typ = listS[1]     #type: 0 - reg, 1 - iteration
        self.name = listS[2]    #name
        self.var = listS[3]     #variable
        self.unit = listS[4]    #unit
        self.exp = listS[5]     #expression
        self.A = listS[6]       #Control variable A
        self.B = listS[7]       #Control variable B
        self.C = listS[8]      #Control variable C
        self.ref = listS[9]    #Reference

        #"Missing" info
        self.lexp =listS[10]     #latex expression
        self.cell = listS[11]   #location in sheet

class imp: #import database variable class
    name = ""   #name
    var = ""    #variable
    path = ""   #path
    A = ""      #Control variable A
    Aunit = ""  #Control variable A unit
    B = ""      #Control variable B
    Bunit = ""  #Control variable B unit
    C = ""      #Control variable C
    Cunit = ""  #Control variable C unit
    D = ""      #Control variable D
    Dunit = ""  #Control variable D unit
    E = ""      #Control variable E
    Eunit = ""  #Control variable E unit
    F = ""      #Control variable F
    Funit = ""  #Control variable F unit
    lexp = ""   #latex equalivent expression
    cell = ""    #location in sheet

    def mCont(self,List,setting):

        tmpList = List

        #Units for import
        tmpList = addList(List,setting[0])
        tmpList = addList(List,setting[1])
        tmpList = addList(List,setting[2])
        tmpList = addList(List,setting[3])
        tmpList = addList(List,setting[4])
        tmpList = addList(List,setting[5])
        
        #Latex expression as it is not generated yet
        tmpList = addList(List,"")

        #Cell Reference
        tmpList = addList(List,"")

        return tmpList
    
    def intCont(self,ListI):
        #Filled in info
        name = ListI[0]   #name
        var = ListI[1]    #variable
        path = ListI[2]   #path
        A = ListI[3]      #Control variable A
        B = ListI[4]      #Control variable B
        C = ListI[5]      #Control variable C
        D = ListI[6]      #Control variable D
        E = ListI[7]      #Control variable E
        F = ListI[8]      #Control variable F

        #"Missing" info
        Aunit = ListI[9]  #Control variable A unit
        Bunit = ListI[10]  #Control variable B unit
        Cunit = ListI[11]  #Control variable C unit
        Dunit = ListI[12]  #Control variable D unit
        Eunit = ListI[13]  #Control variable E unit
        Funit = ListI[14]  #Control variable F unit
        lexp = ListI[15]   #latex equalivent expression
        cell = ListI[16]    #location in sheet
    
class check: #check class
    name = ""         #name
    condition = ""    #condition
    lcond = ""        #latex equalivent condition
    ref = ""          #reference
    cell = ""          #location in sheet

    def mCont(self,List,setting):

        tmpList = List

        #Latex expression as it is not generated yet
        tmpList = addList(List,"")

        #Cell Reference
        tmpList = addList(List,"")

        return tmpList    

    def intCont(self,listC):
        name = listC[0]         #name
        condition = listC[1]    #condition
        ref = listC[3]          #reference

        #"Missing" info
        lcond = listC[3]         #latex equalivent condition
        cell = listC[4]          #location in sheet
    
#Functions
def menuOpt():
    print("================ MENU ================")
    print("Enter the following to options:")
    print("[I] to import list from drive")
    print("[S] to start generating new calculation sheet")
    print("[PV] to print variables list")
    print("[PS] to print sequence list")
    print("[PC] to print check list")
    print("[PI] to print import list")
    print("[E] to exit the program")
    print("[i] to show this menu again")
    print("================ MENU ================")

def addList(List, value):
    if isinstance(List,(list,)):
        temp = List
        temp = temp.append(value)
        return List
    else:
        return [value]

def intLister(num):
    if num == 0:
        return var()
    elif num == 1:
        return seq()
    elif num == 2:
        return imp()
    elif num == 3:
        return check()
    else:
        print("ERROR, Invaild Setting")
        return -1

def intSetting(sheet):
    tmp = [str(sheet['B2'].value)]
    tmp = addList(tmp,str(sheet['B3'].value))
    tmp = addList(tmp,str(sheet['B4'].value))
    tmp = addList(tmp,str(sheet['B5'].value))
    tmp = addList(tmp,str(sheet['B6'].value))
    tmp = addList(tmp,str(sheet['B7'].value))
    tmp = addList(tmp,str(sheet['B8'].value))
    return tmp

def setList(num,vRange,setting):
    tempList = 0
    mcol = vRange.max_column
    mrow = vRange.max_row - 1
    
    for row in vRange.iter_rows(min_row=1, row_offset=1,max_col=mcol,
                                max_row = mrow):
        TList = 0
        for cell in row:
            if(cell.value != None):
                TList = addList(TList,str(cell.value))
            else:
                TList = addList(TList,"")
        tmp = intLister(num)
        TList = tmp.mCont(TList,setting)
        tmp.intCont(TList)

        tempList = addList(tempList,tmp)

    return tempList

def importList(varList,seqList,ckList,iList):
    #Return variables
    vList=0
    sList=0
    cList=0
    iiList=0
    
    #Get path of workbook
    gFile = tkFileDialog.askopenfilename(initialdir = "/",
                                                  title = "Select xlsx file to import",
                                                  filetypes = (("xlsx files","*.xlsx"),
                                                               ("xls files","*.xls"),
                                                               ("all files","*.*")))
    root.destroy()
    if gFile!=None:
        #Load Workbook & Ranges
        WB = xl.load_workbook(gFile)
        vRange = WB['Variables']
        sRange = WB['Sequences']
        iRange = WB['Import']
        cRange = WB['Check']
        stRange = WB['Import Settings']

        #Get Settings
        setting = intSetting(stRange)

        #Load info into list
        vList=setList(0,vRange,setting)
        sList=setList(1,sRange,setting)
        cList=setList(3,cRange,setting)
        iiList = setList(2,iRange,setting)

    return vList,sList,cList,iiList

def generate(varList,seqList,ckList,iList):
    print("Nothing")

#Startup message
def StartupMsg():
    print("======================================")
    print("WELCOME TO CMAKE")
    print("by Regan Lu")
    print("This program allow the user to create")
    print("calculation sheet with ease.")
    print("======================================")
    print("Note: it is recommended to define your") 
    print("variables outside of this program.")
    print("======================================")
    menuOpt();

#Print current Lists functions
def printvarList(varList):
    print("Name | Variable | Value | Unit | Notes | Reference")
    print("==================================================")
    for x in varList:
        print("%s | %s | %s | %s | %s | %s |" %(x.name, x.var, x.value,
                                                x.unit,x.note,x.ref))
    
def printseqList(seqList):
    print("======================================")
def printckList(ckList):
    print("======================================")
def printiList(iList):
    print("======================================")
#Swtich Cases for User Input
def uiInput(Key,varList,seqList,ckList,iList):
    if(Key == 'I'):    #I for Import Lists
        return False,False,True
    elif(Key == 'E'):  #E for Exit
        return False,True,False
    elif(Key == 'S'):  #S for Start
        return True,False,False
    elif(Key == 'PV'): #PV for print Variable List
        printvarList(varList)
        return False,False,False
    elif(Key == 'PS'): #PS for print Sequence List 
        printseqList(seqList)
        return False,False,False
    elif(Key == 'i'):  #i for print this menu
        menuOpt()
        return False,False,False
    else:
        print("Error: Invalid Input")
        return False,False,False

#Main Loop
#Start up
StartupMsg()

#Storage for imports
varList = 0
seqList = 0
ckList = 0
iList = 0

while(exiter == False):
    key = raw_input(">>>")
    starter,exiter,importv = uiInput(key,varList,seqList,ckList,iList)

    if(importv):
        varList,seqList,ckList,iList = importList(varList,seqList,ckList,iList)

    if(starter):
        generate(varList,seqList,ckList,iList)
