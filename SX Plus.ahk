;SX Plus version 0.2

;Set working directory
WorkingDir = \\noble-file\homedc$\duncanp\My Documents\SX Plus\
;WorkingDir = D:\Programs\Autohotkey\SX Plus\

;Read .ini file
IniRead, Printer, %WorkingDir%Settings.ini, section1 , Default Printer
IniRead, SXEmail, %WorkingDir%Settings.ini, section1 , E-Mail Target
IniRead, GWLogin, %WorkingDir%Settings.ini, section1 , GW Login
IniRead, GWPassword, %WorkingDir%Settings.ini, section1 , GW Password
IniRead, SXLogin, %WorkingDir%Settings.ini, section1 , SX Login
IniRead, SXPassword, %WorkingDir%Settings.ini, section1 , SX Password
Loop 13
{
 IniRead, Menu%A_Index%, %WorkingDir%Settings.ini, section1, Menu%A_Index%
}
IniRead, GWOnline, %WorkingDir%Settings.ini, section1 , GW Online
IniRead, SXKeyDelay, %WorkingDir%Settings.ini, section1 , Key Delay
IniRead, SXMouseDelay, %WorkingDir%Settings.ini, section1 , Mouse Delay
IniRead, SXShortcut, %WorkingDir%Settings.ini, section1 , SX Shortcut
IniRead, POReset, %WorkingDir%Settings.ini, section1 , PO Reset Flag
IniRead, WorkingOE, %WorkingDir%Settings.ini, section1 , Current OE
IniRead, WorkingPO, %WorkingDir%Settings.ini, section1 , Current PO

;Set delay so that all clicks and keys are detected
SetKeyDelay %SXKeyDelay%
SetMouseDelay %SXMouseDelay%

SetTitleMatchMode, 1

IfWinNotActive, ahk_class ProMainWin, , ,#2 - Navigator
 WinActivate

;Set the items in the list for the menus
MenuListItem1 = None
MenuListItem2 = OEET
MenuListItem3 = OEIO
MenuListItem4 = POET
MenuListItem5 = POIP
MenuListItem6 = WTET
MenuListItem7 = WTIT
MenuListItem8 = VAET
MenuListItem9 = VAIF
MenuListItem10 = VAES
MenuListItem11 = VAEI
MenuListItem12 = ICIA
MenuListItem13 = ICIP

;Set menu dropdown list options for the menus
Loop 13
{
 MenuList := MenuList . "|" . MenuListItem%A_Index%
}

;Test for PO print setting
If POReset = 1
 FirstPOFlag = 1

;Test for saved Groupwise Online state
If GWOnline = 1
 GWChecked = Checked
Else
 GWChecked =

IfWinNotExist, ahk_class ProMainWin
{
MsgBox, 4, SX. Enterprise Not Detected, Start SX. Enterprise? 
IfMsgBox, Yes
 Gosub StartSX
}

HideFlag = 0
SetFlag = 0
GlassFlag = 0

Zero = 0.000

TWidth := Round(0, 3)
THeight := Round(0, 3)
SLWidth := Round(0, 3)
SLHeight := Round(0, 3)

ControlSpace = 27

WinH = 460
WinW = 93

WinX = 3
WinY = 435
SetWinH := WinH
SetWinW := 220
SetWinX := (WinX + WinW + 6)
SetWinY := WinY

ButtonH := (WinH - 38)
ButtonW := (WinW - 20) 
ButtonY := (WinH - 30)
SetButtonY := (ButtonY - ControlSpace)

;________________Creat GUI Controls__________________
Gui, +alwaysontop -border +toolwindow
Gui, Add, Picture, x0 y0, %WorkingDir%SX Top.bmp
FormatTime, CurrentTime, A_Now, MMM d, yyyy
Gui, Add, Text, x6 y23, %CurrentTime%
Gui, Add, Button, x5 y42 w40, SX
Gui, Add, Button, x49 y42 w40, Email
Gui, Add, Button, x5 y69 w40, Note
Gui, Add, Button, x49 y69 w40, Calc
Gui, Add, Button, x5 y96 w40, Excel
Gui, Add, Button, x49 y96 w40, Word
Gui, Add, Button, x5 y123 w40, VA
Gui, Add, Button, x49 y123 w40, Glass
Gui, Add, Text, x10 y154, OE#:
Gui, Add, Edit, x39 y151 w50 +Right +Limit7 vCurrentOE
Gui, Add, Button, x5 y177 w84, Print OE
Gui, Add, Button, x5 y204 w84, Email OE
Gui, Add, Text, x10 y236, PO#:
Gui, Add, Edit, x39 y232 w50 +Right +Limit7 vCurrentPO
Gui, Add, Button, x5 y257 w84, Print PO
Gui, Add, Button, x5 y284 w84, Add Note
Gui, Add, Button, x5 y378 w84, Hide
Gui, Add, Button, x5 y405 w84, Settings
Gui, Add, Button, x5 y432 w84, Exit
GuiControl,, CurrentOE, %WorkingOE%
GuiControl,, CurrentPO, %WorkingPO%
Gui, Show, w%WinW% h%WinH% x%WinX% y%WinY%, MainWin

;Settings Window
Gui, 2: +alwaysontop -border +toolwindow
Gui, 2: Add, Picture, x0 y0, %WorkingDir%Settings Top.bmp
Gui, 2: Add, Text, x5 y25, Printer:
Gui, 2: Add, Edit, x60 y22 w50 +Limit5 vPrinter, %Printer%
Gui, 2: Add, Text, x5 y52, SX Email:
Gui, 2: Add, Edit, x60 y49 w155 +Limit vSXEmail, %SXEmail%
Gui, 2: Add, Text, x5 y79, GW Login:
Gui, 2: Add, Edit, x60 y76 w155 +Limit vGWLogin, %GWLogin%
Gui, 2: Add, Text, x5 y106, GW Pass:
Gui, 2: Add, Edit, x60 y103 w155 +Limit +Password vGWPassword, %GWPassword%
Gui, 2: Add, Checkbox, %GWChecked% vGWOnline, Launch GroupWise Online 
Gui, 2: Add, Text, x5 y153, SX Login:
Gui, 2: Add, Edit, x60 y150 w50 +Limit4 vSXLogin, %SXLogin%
Gui, 2: Add, Text, x5 y180, SX Pass:
Gui, 2: Add, Edit, x60 y177 w155 +Limit +Password vSXPassword, %SXPassword%
Gui, 2: Add, Text, x5, Launch additional windows at startup:

LoopX := 5
LoopY := 222
LoopVar := 1
Loop 12
{
 Gui, 2: Add, DropDownList, x%LoopX% y%LoopY% w65 vMenuDropDown%A_Index%
 GuiControl, 2:, MenuDropDown%A_Index%, %MenuList%
 Loop 12
 {
  If Menu%LoopVar% = % MenuListItem%A_Index%
   GuiControl, 2: Choose, MenuDropDown%LoopVar%, %A_Index%
  If Menu%LoopVar% = % MenuListItem13
   GuiControl, 2: Choose, MenuDropDown%LoopVar%, 13
 }   
 LoopY := LoopY + 25
 LoopVar++
 If LoopVar = 9
 {
  LoopX := 151
  LoopY := 222
 }
 Else If LoopVar = 5
 {
  LoopX := 78
  LoopY := 222
 }
}

Gui, 2: Add, Text, x5 y326, Key Delay:
Gui, 2: Add, Edit, x69 y323 w35 +Limit3 vSXKeyDelay, %SXKeyDelay%
Gui, 2: Add, Text, x115 y326, Click Delay:
Gui, 2: Add, Edit, x179 y323 w35 +Limit3 vSXMouseDelay, %SXMouseDelay%
Gui, 2: Add, Text, x5 y353, SX Shortcut:
Gui, 2: Add, Edit, x69 y350 w145 +Limit100 r1 vSXShortcut, %SXShortcut%
Gui, 2: Add, Button, x5 y378 w211, Help
Gui, 2: Add, Button, x5 y405 w211, About
Gui, 2: Add, Button, x5 y432 w211, Save && Exit
Gui, 2: Show, w%SetWinW% h%SetWinH% x%SetWinX% y%SetWinY%, SetWin
Gui, 2: Hide

Gui, 3: +alwaysontop -border +toolwindow
Gui, 3: Add, Picture, x0 y0, %WorkingDir%Glass Top.bmp
Gui, 3: Add, Text, x10 y23, Unit Type:
Gui, 3: Add, DropDownList, vUnitType, Single||Double|SD|SDD|SDS|SDDS
Gui, 3: Add, CheckBox, Checked vTransom, Transom
Gui, 3: Add, Radio, Group x73 y90 vInterior, Interior
Gui, 3: Add, Radio, Checked x10 y90 vExterior, Exterior
Gui, 3: Add, Radio, Group Checked x10 y110 vInswing, Inswing
Gui, 3: Add, Radio, x73 y110 vOutswing, Outswing
Gui, 3: Add, Text, x10 y135, Slab Width:
Gui, 3: Add, Edit, x78 y132 w50 +Right +Limit7 vDoorWidth, 0
Gui, 3: Add, Text, x10 y162, Slab Height:
Gui, 3: Add, Edit, x78 y159 w50 +Right +Limit7 vDoorHeight, 0
Gui, 3: Add, Text, x10 y189, RO Width:
Gui, 3: Add, Edit, x78 y186 w50 +Right +Limit7 vROWidth, 0
Gui, 3: Add, Text, x10 y216, RO Height:
Gui, 3: Add, Edit, x78 y213 w50 +Right +Limit7 vROHeight, 0
Gui, 3: Add, Button, x10 y243 w118, Calculate
Gui, 3: Add, Text, x4 y266, ______________________
Gui, 3: Add, Text, x10 y287, Trans Width:
Gui, 3: Add, Text, x82 y287 w50 vTWidth, %TWidth%
Gui, 3: Add, Text, x10 y314, Trans Height:
Gui, 3: Add, Text, x82 y314 w50 vTHeight, %THeight%
Gui, 3: Add, Text, x10 y341, S/L Width:
Gui, 3: Add, Text, x82 y341 w50 vSLWidth, %SLWidth%
Gui, 3: Add, Text, x10 y368, S/L Height:
Gui, 3: Add, Text, x82 y368 w50 vSLHeight, %SLHeight%
Gui, 3: Add, Button, x5 y405 w130, Reset
Gui, 3: Add, Button, x5 y432 w130, Exit 
Gui, 3: Show, w140 h%SetWinH% x%SetWinX% y%SetWinY%, Glass Calculator
Gui, 3: Hide

Gui, 4: +alwaysontop -sysmenu +owner1 +toolwindow
Gui, 4: Add, Text, x5 y8, Subject:
Gui, 4: Add, Edit, x47 y5 w142 +Right +Limit vNoteOEName
Gui, 4: Add, Button, x5 y31 w184 default, Add Note
Gui, 4: Show, w193 h58, PO Note Info
Gui, 4: Hide

Gui, 5: +alwaysontop -sysmenu +owner1 +toolwindow
Gui, 5: Add, Text, x5 y8, Subject:
Gui, 5: Add, Edit, x47 y5 w142 +Right +Limit vEmailOEName
Gui, 5: Add, Button, x5 y31 w184 default, Send Email
Gui, 5: Show, w193 h58, OE Email Info
Gui, 5: Hide

Gui, 6: +alwaysontop -sysmenu +owner1 +toolwindow
Gui, 6: Add, Text, x15 y13, Please define labour type.
Gui, 6: Add, Button, x15 y42 w78 h23 +Default, Interior
Gui, 6: Add, Button, x100 y42 w78 h23, Exterior
Gui, 6: Show, w193 h78, Labour Info
Gui, 6: Hide

;_______________ This is where the program waits _________________
MainProgram:

#1:: GoSub ButtonSX
#2:: GoSub ButtonEmail
#3:: GoSub ButtonNote
#4:: Gosub ButtonCalc
#5:: Gosub ButtonExcel
#6:: Gosub ButtonWord
#7:: Gosub ButtonVA
#8:: Gosub ButtonGlass

;Alt Click to copy text to clipboard
!LButton:: 
{
IfWinActive, Sales Order Document Print
 {
  MouseGetPos, , , id, TempControl
  ControlGet, TempOENumber, Line, 1, %TempControl%, Sales Order Document Print
  CurrentOENumber = %TempOENumber%
  GuiControl,, CurrentOE, %CurrentOENumber%
  Exit
 }
IfWinActive, Purchase Order Document Print
 {
  MouseGetPos, , , id, TempControl
  ControlGet, TempPONumber, Line, 1, %TempControl%, Purchase Order Document Print
  CurrentPONumber = %TempPONumber%
  GuiControl,, CurrentPO, %CurrentPONumber%
  Exit
 }
IfWinActive, #2 - OE
 {
  MouseGetPos, , , id, TempControl
  ControlGet, TempOENumber, Line, 1, %TempControl%, #2 - OE
  CurrentOENumber = %TempOENumber%
  GuiControl,, CurrentOE, %CurrentOENumber%
  Exit
 }
IfWinActive, #2 - PO
 {
  MouseGetPos, , , id, TempControl
  ControlGet, TempPONumber, Line, 1, %TempControl%, #2 - PO
  CurrentPONumber = %TempPONumber%
  GuiControl,, CurrentPO, %CurrentPONumber%
  Exit
 }
}
Send +{Click}
Exit

+Lbutton::
{
Gui 1: Submit, Nohide
IfWinActive, Sales Order Document Print
 {
 Clipboard = %CurrentOE%
 Click
 Send {End}
 Send +{Home}
 Send ^v
 Exit
 }
IfWinActive, Purchase Order Document Print
 {
 Clipboard = %CurrentPO%
 Click
 Send {End}
 Send +{Home}
 Send ^v
 Exit
 }
IfWinActive, #2 - OE
 {
 Clipboard = %CurrentOE%
 Click
 Send {End}
 Send +{Home}
 Send ^v
 Exit
 }
IfWinActive, #2 - PO
 {
 Clipboard = %CurrentPO%
 Click
 Send {End}
 Send +{Home}
 Send ^v
 Exit
 }
}
Send +{Click}
Exit

#`::  ; Hide/Show the toolbar
Gosub HideShow
Return

Pause

;______________________Buttons________________________
ButtonExit:
Gui, Destroy
Gui, 2: Destroy
Gui, 3: Destroy
ExitApp

ButtonSettings: ;Hide/Show settings window
If SetFlag = 0
{
 Gui, 3: Hide
 If GlassFlag = 1
  GlassFlag = 0
 Gui, 2: Show
 SetFlag = 1
 Return
}
If SetFlag = 1
{
 Gui, 2: Hide
 SetFlag = 0
 Return
}

HideShow:
If HideFlag = 0
{
 Gui, Hide
 Gui, 2: Hide
 Gui, 3: Hide
 HideFlag = 1
 Return
}
If HideFlag = 1
{
 Gui, Restore
 HideFlag = 0
 If SetFlag = 1
  Gui, 2: Show
 If GlassFlag = 1
  Gui, 3: Show
 Return
}

2ButtonHelp:
Msgbox , , SX. Plus Help, SX. Plus is intented to run with the SX. Enterprise window maximized. `n`nToday's date is displayed on the top of the main window.`n`nThe first eight shortcuts may be accessed by holding the Windows Key and pressing 1-8. `n1. Launch or activate SX. Enterprise.`n2. Launch or activate GroupWise. GroupWise login information can be changed in the settings. `n3. Launch or activate Notepad. `n4. Launch or activate Calculator. `n5. Launch or active Openoffice Calc. `n6. Launch or activate Openoffice Writer. `n7. Add sections to a VA. `n8. Open the glass calculator tab. `n`nThe OE# and PO# boxes are used to print or e-mail orders. `nHolding ALT and left-clicking on the OE# or PO# in SX. Enterprise will load the number for later use.`n`nThe Email OE button will send a copy of the OE to the SX e-mail address defined in the settings. `n`nThe Add Note button will write a global note to the PO defined in the box above. `n`nThe Hide button will hide SX. Plus. Holding windows key and pressing tilde (the key left of 1) will `nshow or hide SX. Enterprise. `n`nThe Settings button opens a tab in which various preferences may be changed. This includes the`ndefault printer and login information, the key and click delays (in milliseconds) and the target of the`nSx. Enterprise shortcut on the desktop. `n`nIt's also possible to automatically load windows when booting SX. Enterprise by using the drop-down`nmenus. The loading order for the windows is top-to-bottom, left-to-right.
Return

2ButtonAbout:
Msgbox , , About SX. Plus, SX. Plus v0.2`nby Duncan Priebe`n`nSX. Plus is designed to enhance your experience `nwith SX. Enterprise by using simple shortcuts.`n`nSupport & Information: duncanpriebe@gmail.com
Return

2ButtonSaveExit:
Gui 1: Submit, Nohide
Gui, 2: Submit
IniWrite, %Printer%, %WorkingDir%Settings.ini, section1 , Default Printer
IniWrite, %SXEmail%, %WorkingDir%Settings.ini, section1 , E-Mail Target
IniWrite, %GWLogin%, %WorkingDir%Settings.ini, section1 , GW Login
IniWrite, %GWPassword%, %WorkingDir%Settings.ini, section1 , GW Password
IniWrite, %SXLogin%, %WorkingDir%Settings.ini, section1 , SX Login
IniWrite, %SXPassword%, %WorkingDir%Settings.ini, section1 , SX Password
Loop 13
{
 IniWrite, % MenuDropDown%A_Index%, %WorkingDir%Settings.ini, section1, Menu%A_Index%
}
IniWrite, %GWOnline%, %WorkingDir%Settings.ini, section1 , GW Online
IniWrite, %SXKeyDelay%, %WorkingDir%Settings.ini, section1 , Key Delay
IniWrite, %SXMouseDelay%, %WorkingDir%Settings.ini, section1 , Mouse Delay
IniWrite, %SXShortcut%, %WorkingDir%Settings.ini, section1 , SX Shortcut
If FirstPOFlag = 1
 POReset = 1
else
 POReset = 0
IniWrite, %POReset%, %WorkingDir%Settings.ini, section1 , PO Reset Flag
IniWrite, %CurrentOE%, %WorkingDir%Settings.ini, section1 , Current OE
IniWrite, %CurrentPO%, %WorkingDir%Settings.ini, section1 , Current PO
Gui, 2: Hide
SetFlag = 0
Reload
Return

ButtonSX:
IfWinExist, ahk_class ProMainWin,,, #2 - OE Entry Transactions
 WinActivate
IfWinNotExist, ahk_class ProMainWin
GoSub StartSX
Return

StartSX:
FirstPOFlag = 0
Run %SXShortcut%
;Run C:\Documents and Settings\All Users\Desktop\Sxe DnC GUI LIVE
WinWait Welcome to SX.enterprise
WinActivate
Send {tab}
Send %SXLogin%
Send {tab}%SXPassword%
Send {Enter}
WinWait #2 - Navigator
WinActivate
WinWait #2 - Report Viewer
WinActivate
WinClose #2 - Report Viewer
WinWait #2 - Navigator
WinActivate
SendEvent {click, 674, 497, down}{click 522, 361, up}

WindowFlag = 0

Loop 12
{
 If Menu%A_Index% != None  
 { 
  Sleep, 500
  If WindowFlag = 1
   Click 600, 50
  Send % Menu%A_Index%
  Send {Enter}
  Sleep, 800
  WinWait, , Forward
  WinActivate
  If WindowFlag = 0
   Gosub SetFirstWindow
 }  
}
Click 79,964
 
Return

SetFirstWindow:
WinMaximize
WindowFlag = 1
Return
 
;Execute when Email button is pressed
ButtonEmail:
If GWOnline = 1
{
 IfWinExist, , Novell WebAccess
 {  
  WinActivate
 }
 Else
 {
 Run http://76.77.68.74:8080/gw/webacc
 WinWait, , Novell WebAccess
 WinActivate
 Sleep, 100
 Send %GWLogin%
 Send {Tab}
 Send %GWPassword%
 Send {Enter}
 }
}
Else
 {
  IfWinExist, ahk_class OFCalView
 {
   WinActivate
 }
 Else
 {
   Run D:\Program Files (x86)Novell\GroupWise\grpwise.exe
   WinWait GroupWise Password
   WinActivate
   Send %GWPassword%
   Send {Enter}
   WinWait Novell GroupWise - Mailbox
   WinMaximize
 }
}
Exit

;Execute when Note button is pressed
ButtonNote:
IfWinExist, ahk_class Notepad
 WinActivate
Else
{
 Run Notepad
 WinWait, Untitled - Notepad
 WinActivate
}
Return

;Execute when Calc button is pressed
ButtonCalc:
IfWinExist, ahk_class SciCalc
 WinActivate
Else
{
 Run Calc
 WinWait Calculator
 WinActivate
}
Return

ButtonGlass:
If GlassFlag = 0
{
 Gui, 2: Hide
 If SetFlag = 1
  SetFlag = 0
 Gui, 3: Show
 GlassFlag = 1
 Return
}
If GlassFlag = 1
{
 Gui, 3: Hide
 GlassFlag = 0
 Return
}

;Execute when Excel button is pressed
ButtonExcel:
IfWinExist, Untitled 1 - OpenOffice.org Calc
 WinActivate
Else
{
 Run D:\Program Files (x86)\OpenOffice.org 3\program\scalc.exe
 WinWait Untitled 1 - OpenOffice.org Calc
 WinActivate
}
Return

;Execute when Word button is pressed
ButtonWord:
IfWinExist, Untitled 1 - OpenOffice.org Writer
 WinActivate
Else
{
 Run D:\Program Files (x86)\OpenOffice.org 3\program\swriter.exe
 WinWait Untitled 1 - OpenOffice.org Writer
 WinActivate
}
Return

ButtonVA:
IfWinNotExist, , Add New Section
 {
  MsgBox, , Error - Incorrect Step, Use command at correct step.
  Exit
 }
 WinWait, , Add New Section
 WinActivate
 Click 983, 891 
 WinWait VA Entry Transactions - Add VA Order, , 1
 IfWinExist, SX.enterprise Error Number (6569)
 { 
  WinActivate
  Click 184, 92
  Exit
 }
WinActivate
Send i{Enter}i{Enter} 
WinWait Inventory Components Extended Information For Seq#: 1
WinActivate
Send {Tab}{Tab}ZDOR{Enter}{Enter}
WinActivate
Send {Enter}
WinWait VA Entry Transactions - Add VA Order
WinActivate
Send ii{Enter}f{Enter}
WinWait Internal Process Extended Information For Seq#: 2
WinActivate
Send ZDOR{Tab}f{Enter}{Enter}
MsgBox, 4, #2 - VA Entry Transactions, Add labour to VA Fabrication?
IfMsgBox, Yes
 {
 Click 285, 204
 Click 384, 137
 Gui, 6: Show
 }
Exit

ButtonPrintOE:
IfWinNotExist, #2 - OE Entry Transactions
 {
  MsgBox, , Error - Incorrect Step, Use command at correct step.
  Exit
 }
Gui 1: Submit, Nohide
WinWait  #2 - OE Entry Transactions
WinActivate
Send {Alt}{F}{P}
WinWait Sales Order Document Print
If CurrentOE = 
 Exit
Send %CurrentOE%{Tab}{Tab}Y{Tab}P
Send {Tab}
Send %Printer%
Send +{Tab}+{Tab}
Send {Down}
Send N
Send {Tab}{Tab}{Tab}
Send {Enter}
Exit

ButtonEmailOE:
IfWinNotExist, #2 - OE Entry Transactions
 {
  MsgBox, , Error - Incorrect Step, Use command at correct step.
  Exit
 }
Gui 1: Submit, Nohide
StringLen, TestString, CurrentOE
If TestString < 6
{
 MsgBox, ,  Error - Insufficient Data, Please enter all required fields.
 Exit
}
Gui, 5: Show
Exit

5ButtonSendEmail:
Gui, 5: Submit
WinWait  #2 - OE Entry Transactions
WinActivate
Send {Alt}{F}{P}
WinWait Sales Order Document Print
Send %CurrentOE%{Tab}{Tab}Y{Tab}E{Tab}{Tab}
Send +{End}{BS}
Send %SXEmail%{Tab}
Send %EmailOEName% %CurrentOE%
Send {Tab}
Send {Enter}
WinWait #2 - OE Entry Transactions
WinActivate
MsgBox, 4, #2 - OE Entry Transactions, Print order?
IfMsgBox, Yes
 {
  Gosub ButtonPrintOE
 }
Exit

ButtonPrintPO:
IfWinNotExist, #2 - PO Entry Transactions
 {
  MsgBox, , Error - Incorrect Step, Use command at correct step.
  Exit
 }
Gui 1: Submit, Nohide
WinWait  #2 - PO Entry Transactions
WinActivate
Send {Alt}{F}{P}
WinWait Purchase Order Document Print
WinActivate
If CurrentPO = 
 Exit
Send %CurrentPO%
Send {Tab}{Tab}
If FirstPOFlag = 0
{
 Send {Space}
 FirstPOFlag = 1
}
Send {Tab}{Tab}P{Tab}
Send %Printer%
Send {Tab}{Tab}
Send {Enter}
Exit

ButtonAddNote:
IfWinNotExist, #2 - PO Entry Transactions
{
 MsgBox, , Error - Incorrect Step, Use command at correct step.
 Exit
}
Gui 1: Submit, Nohide
StringLen, TestString, CurrentOE
If TestString < 6
{
 MsgBox, ,  Error - Insufficient Data, Please enter all required fields.
 Exit
}
StringLen, TestString, CurrentPO
If TestString < 7
{
 MsgBox, ,  Error - Insufficient Data, Please enter all required fields.
 Exit
}
Gui, 4: Show
Exit

4ButtonAddNote:
Gui, 4: Submit
WinWait  #2 - PO Entry Transactions
WinActivate
Click 132, 106
Send {Tab}{Tab}{Tab}
Send %CurrentPO%
Click 181, 48
WinWait, Display Notes
WinActivate
Send !2
Sleep, 150
Send B
Sleep, 150
Send {Enter}
Sleep, 150
Send %NoteOEName% %CurrentOE%
Click 135, 454
Click 512, 485
Send {Enter}
Send {Enter}
WinWait, #2 - PO Entry Transactions
WinActivate
MsgBox, 4, #2 - PO Entry Transactions, Print purchase order?
IfMsgBox, Yes
 {
  Gosub ButtonPrintPO
 }
Exit

ButtonHide:
Gosub HideShow
Exit

;________________________Calculate Glass_________________________

3ButtonCalculate:
Gui 3: Submit, Nohide
If UnitType = Single
{
 SLWidth := Round(0, 3)
 SLHeight := Round(0, 3)
 If DoorWidth = 0
  Gosub ErrorReset 
 If DoorHeight = 0
  Gosub ErrorReset 
 If Transom = 0 
  Gosub ErrorReset
 If ROHeight = 0
  Gosub ErrorReset
 If Interior = 1
  {
  THeight := Round((ROHeight - DoorHeight - 4), 3)
  TWidth := Round(DoorWidth, 3)
 }
 If Exterior = 1
 {
  If Inswing = 1
   THeight := Round((ROHeight - DoorHeight - 5), 3)
  If Outswing = 1
   THeight := Round((ROHeight - DoorHeight - 4), 3)
  TWidth := Round(DoorWidth, 3)
 }
}

If UnitType = Double
{
 SLWidth := Round(0, 3)
 SLHeight := Round(0, 3)
 If DoorWidth = 0
  Gosub ErrorReset 
 If DoorHeight = 0
  Gosub ErrorReset 
 If Transom = 0 
  Gosub ErrorReset
 If ROHeight = 0
  Gosub ErrorReset
 If Interior = 1
  {
   THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((DoorWidth *2), 3)
  }
 If Exterior = 1
 {
  If Inswing = 1
   THeight := Round((ROHeight - DoorHeight - 5), 3)
  If Outswing = 1
   THeight := Round((ROHeight - DoorHeight - 4), 3)
  TWidth := Round((DoorWidth *2) + (0.875), 3)
 }
}

If UnitType = SD
{
  If DoorWidth = 0
   Gosub ErrorReset 
  If DoorHeight = 0
   Gosub ErrorReset 
  If ROWidth = 0
   Gosub ErrorReset
  If Transom = 0
 {
  TWidth := Round(0, 3)
  THeight := Round(0, 3)
 }
 If Transom = 1 
 {
  If ROHeight = 0
   Gosub ErrorReset
  If Interior = 1
  {
   THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
   SLHeight := Round((DoorHeight), 3)
   SLWidth := Round((ROWidth - DoorWidth) - (3.75), 3)
  }
  If Exterior = 1
  {
   If Inswing = 1
    THeight := Round((ROHeight - DoorHeight - 5), 3)
   If Outswing = 1
    THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
  }
 }
 If Inswing = 1
  SLHeight := Round((DoorHeight) + (0.5), 3)
 If Outswing = 1
  SLHeight := Round((DoorHeight), 3)
 SLWidth := Round((ROWidth - DoorWidth) - (3.75), 3)
}

If UnitType = SDD
{
  If DoorWidth = 0
   Gosub ErrorReset 
  If DoorHeight = 0
   Gosub ErrorReset 
  If ROWidth = 0
   Gosub ErrorReset
  If Transom = 0
 {
  TWidth := Round(0, 3)
  THeight := Round(0, 3)
 }
 If Transom = 1 
 {
  If ROHeight = 0
   Gosub ErrorReset
  If Interior = 1
  {
   THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
   SLHeight := Round((DoorHeight), 3)
   SLWidth := Round(ROWidth - (DoorWidth *2) - (4.625), 3)
  }
  If Exterior = 1
  {
   If Inswing = 1
    THeight := Round((ROHeight - DoorHeight - 5), 3)
   If Outswing = 1
    THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
  }
 }
 If Inswing = 1
  SLHeight := Round((DoorHeight) + (0.5), 3)
 If Outswing = 1
  SLHeight := Round((DoorHeight), 3)
 SLWidth := Round(ROWidth - (DoorWidth *2) - (4.625), 3)
}

If UnitType = SDS
{
  If DoorWidth = 0
   Gosub ErrorReset 
  If DoorHeight = 0
   Gosub ErrorReset 
  If ROWidth = 0
   Gosub ErrorReset
  If Transom = 0
 {
  TWidth := Round(0, 3)
  THeight := Round(0, 3)
 }
 If Transom = 1 
 {
  If ROHeight = 0
   Gosub ErrorReset
  If Interior = 1
  {
   THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
   SLHeight := Round((DoorHeight), 3)
   SLWidth := Round(((ROWidth - DoorWidth - 5.25) /2), 3)
  }
  If Exterior = 1
  {
   If Inswing = 1
    THeight := Round((ROHeight - DoorHeight - 5), 3)
   If Outswing = 1
    THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
  }
 }
 If Inswing = 1
  SLHeight := Round((DoorHeight) + (0.5), 3)
 If Outswing = 1
  SLHeight := Round((DoorHeight), 3)
 SLWidth := Round(((ROWidth - DoorWidth - 5.25) /2), 3)
}

If UnitType = SDDS
{
  If DoorWidth = 0
   Gosub ErrorReset 
  If DoorHeight = 0
   Gosub ErrorReset 
  If ROWidth = 0
   Gosub ErrorReset
  If Transom = 0
 {
  TWidth := Round(0, 3)
  THeight := Round(0, 3)
 }
 If Transom = 1 
 {
  If ROHeight = 0
   Gosub ErrorReset
  If Interior = 1
  {
   THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
   SLHeight := Round((DoorHeight), 3)
   SLWidth := Round(((ROWidth - (DoorWidth *2) - 6) /2), 3)
  }
  If Exterior = 1
  {
   If Inswing = 1
    THeight := Round((ROHeight - DoorHeight - 5), 3)
   If Outswing = 1
    THeight := Round((ROHeight - DoorHeight - 4), 3)
   TWidth := Round((ROWidth) - (2.125), 3)
  }
 }
 If Inswing = 1
  SLHeight := Round((DoorHeight) + (0.5), 3)
 If Outswing = 1
  SLHeight := Round((DoorHeight), 3)
 SLWidth := Round(((ROWidth - (DoorWidth *2) - 6) /2), 3)
}

GuiControl,, TWidth, %TWidth%
GuiControl,, THeight, %THeight%
GuiControl,, SLWidth, %SLWidth%
GuiControl,, SLHeight, %SLHeight%
Return

; Reset values upon incorrect data input
ErrorReset:
TWidth := Round(0, 3)
THeight := Round(0, 3)
SLWidth := Round(0, 3)
SLHeight := Round(0, 3)
GuiControl,, TWidth, %TWidth%
GuiControl,, THeight, %THeight%
GuiControl,, SLWidth, %SLWidth%
GuiControl,, SLHeight, %SLHeight%
MsgBox, ,  Error - Insufficient Data, Please enter all required fields.
Exit

3ButtonReset:
TWidth = %Zero%
THeight = %Zero%
SLWidth = %Zero%
SLHeight = %Zero%
GuiControl,, TWidth, %TWidth%
GuiControl,, THeight, %THeight%
GuiControl,, SLWidth, %SLWidth%
GuiControl,, SLHeight, %SLHeight%
GuiControl,, DoorWidth, 0
GuiControl,, DoorHeight, 0
GuiControl,, ROWidth, 0
GuiControl,, ROHeight, 0
GuiControl,, Transom, 1
GuiControl,, Inswing, 1
GuiControl,, Exterior, 1
GuiControl,, UnitType, |Single||Double|SD|SDD|SDS|SDDS
Return

3ButtonExit:
Gui, 3: Hide
GlassFlag = 0
Return

6ButtonInterior:
Gui, 6: Hide
Sleep 100
WinActivate, , Add New Section
Click 971, 901
WinWait VA Entry Transactions - Add VA Line
WinActivate
Send LABOUR-DOOR
Send {Tab 4}
Send {Enter}
WinWait VA Entry Transactions - Labor Product (Estimated/Actual)
WinActivate
Send {Tab}
Send %SXLogin%
Send {Tab 5}E{Tab 2}{Space}
Send +{Tab}
Send +{Tab}
Send +{Tab}
Send +{Tab}
Return

6ButtonExterior:
Gui, 6: Hide
Sleep 100
WinActivate, , Add New Section
Click 971, 901
WinWait VA Entry Transactions - Add VA Line
WinActivate
Send LABOUR-DOORX
Send {Tab 4}
Send {Enter}
WinWait VA Entry Transactions - Labor Product (Estimated/Actual)
WinActivate
Send {Tab}
Send %SXLogin%
Send {Tab 5}E{Tab 2}{Space}
Send +{Tab}
Send +{Tab}
Send +{Tab}
Send +{Tab}
Return