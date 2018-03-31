#Page Configuration
#------------------ 

#[Size]
SizeX = 8
SizeY = 44

#[Locations]
Logo = 'A1'
Project = 'F1'
Title = 'F2'
Name = 'F3'
Date = 'F4'
Page = 'H4'

#[Variables]
rgVarN = [1,2]
cVarL = 3
cValue = 4
cUnit = 5

#[Formula]
#Control Cell is 0,0
colForm = 'D'   #Column of cell reference (control cell)
cRef = 4        #Reference col from cell ref
rRef = -2       #Reference col from cell ref
rEqt = -3       #Eqn Title row from  cell ref
cEqt = -3       #Eqn Title col from cell ref
rgEqt = [-3,0]  #Eqn Title col range from cell ref
cVar = -1       #Eqn Var col from cell ref
rVar = -3       #Eqn Var row from cell ref
rgVar = [-3,-2] #Eqn Var range from cell ref
rForm = -2      #Eqn formula row from cell ref
rOver = -1      #Eqn override row from cell ref
cNote = -2      #Eqn Notes col from cell ref
rNote = 1       #Eqn Notes row from cell ref
fSize = 5       #Eqn size of block

#[If-Func]
ifSize = 3
cift = -3
rift = -1
cifV = -1
rifV = 0

#[Lookup]

#[Import Table]
colTable = 'A'  #Column of cell import table

#[Colour-Coding]
cColour = 'ff9999'
cColR = 'B37'
cDesR = 'C37:E37'
cDes = 'Control Cell: Do not modify.'

oColour = 'ffc04c'
oColR = 'B38'
oDesR = 'C38:E38'
oDes = 'Override Cell: Allows Designer Input'

fColour = '99cc99'
fColR = 'B39'
fDesR = 'C39:E39'
fDes = 'Formula Cell: Calculates based on Equation'

#[Picture]
Logo = 'BML-Logo-New.png'
TitlePic = 'BML-Logo-New.png'

#[first-page]
fTitle = 'B6'
tRange = 'B6:E6'
tPict = 'A10'

#[Others]
spacing = 1
tblock = 7
cpyRange = 'A1:H44'
