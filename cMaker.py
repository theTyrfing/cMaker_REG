#cMaker - Calculation Maker
#by Regan Lu

#Libaraies
import numpy as np
import  
import
import

#Variables
loaded = False
exiter = False
starter = False

#Storage for imports

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

def importVar():
    

def generate(varList,seqList,ckList,iList):
    

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
    
def printseqList(seqList):
    
def printckList(ckList):
    
def printiList(iList):
    
#Swtich Cases for User Input
def uiInput(Key):
    if(Key == 'I'):    #I for Import Lists
        importVar()
        return False,False
    elif(Key == 'E'):  #E for Exit
        return False,True
    elif(Key == 'S'):  #S for Start
        return True,False
    elif(Key == 'PV'): #PV for print Variable List
        printvarList()
        return False,False
    elif(Key == 'PS'): #PS for print Sequence List 
        printseqList()
        return False,False
    elif(Key == 'i'):  #i for print this menu
        menuOpt()
        return False,False
    else:
        #do nothing
        
#Main Loop

#Start up
StartupMsg()

while(exiter == False):
    key =  
   [starter,exiter] = uiInput(key)
    