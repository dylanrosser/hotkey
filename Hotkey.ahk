;Author: Dylan Rosser 2018
;Modified April 2021

; This is a script that will run or switch to certain applications. It will also launch common directories like the C drive.
;Add as many of these as you want to a .ahk file, and set that to be run at startup.

;   # - Windows Key
;   + - Shift Key
;   ^ - Control Key
;   ! - Alt Key

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance
#WinActivateForce
DetectHiddenWindows, on
SetTitleMatchMode, 2 ;These two were added becuase I  was having trouble Winactivating google chrome.
SetTitleMatchMode, Fast


;-----------------------------------------------------------------------------
;Win+Shift+q = tells the class of the active window (used for debugging)
+#q::
WinGetClass, class, A
MsgBox, The active window's class is "%class%".
return
;-----------------------------------------------------------------------------
;~ ;Win+Shift+D = Remind Everyone that Dylan is Awesome
;~ +#D::
;~ SetTimer, ChangeButtonNames, 50 
;~ MsgBox, 52, Awesomeness Alert!, Dylan is Awesome
;~ return 
;~ ChangeButtonNames: 
;~ IfWinNotExist, Awesomeness Alert!
    ;~ return  ; Keep waiting.
;~ SetTimer, ChangeButtonNames, Off 
;~ WinActivate 
;~ ControlSetText, Button1, &I Agree 
;~ ControlSetText, Button2, &Definately 
;~ return

;Win+M = Minimize the active Window
#m::WinMinimize, A
return
;-----------------------------------------------------------------------------
;CTRL+Q = Quit the Active Application
^q::WinClose, A
return
;-----------------------------------------------------------------------------
;Win+T Show Terminal
#t::
if WinExist("Windows PowerShell")
    {
        WinActivate
        return
    }
else
    ;{
        Run, powershell -NoExit -Command "cd $Env:REPOS"
        return
    ;}
;-----------------------------------------------------------------------------
;Win+V Show VS Code
#v::
if WinExist("Visual Studio Code")
    {
        WinActivate
        return
    }
else
    {
        Run, code
        Sleep, 3000
        WinClose, C:\WINDOWS\system32\cmd.exe
        return
    }
;-----------------------------------------------------------------------------
;Win+H = Run/Switch to SciTE Editor
#h::
if WinExist("ahk_class SciTEWindow")
    {
        WinActivate
        return
    }
else
    ;{
        Run "C:\Program Files\AutoHotkey\SciTE\SciTE.exe"
       return
    ;}
;-----------------------------------------------------------------------------
; Win+Shift+H = Open Hotkey Folder
#+h::
Run "C:\Users\Dylan\Documents\AutoHotkey\hotkey-master"
return
;-----------------------------------------------------------------------------
;Win+E = Open Explorer if not already open
#E::
if WinExist("ahk_class CabinetWClass")
    {
        WinActivate
        return
    }
else
    {
     Run, Explorer   
        return
    }
;-----------------------------------------------------------------------------
;Win+Shift+X = Open an active excel sheet or launch excel
#+x::
if WinExist("ahk_class XLMAIN")
    {
        
        xlApp := ComObjActive("Excel.Application")
        WinActivate, % "ahk_id " xlApp.Hwnd
        return
    }
else
    {
        Run Excel
        return
    }
;-----------------------------------------------------------------------------
;Win+Shift+W= Open an active word file or launch MS Word
#+w::
if WinExist("ahk_class OpusApp")
    {
        WinActivate
        return
    }
else
    {
        Run WinWord
        return
    }
;-----------------------------------------------------------------------------
;Alt+C = Open / Switch to calculator
!c::

	If WinExist("Calculator ahk_class ApplicationFrameWindow")
		WinActivate
	else
		Run, calc
    return

;-----------------------------------------------------------------------------
;~ ; Win+Shift+O = Open Outlook
;~ #+o::
    ;~ if WinExist("ahk_class rctrl_renwnd32")
    ;~ {
        ;~ WinActivate
        ;~ return
    ;~ }
    ;~ else 
    ;~ {
        ;~ Run "C:\Program Files (x86)\Microsoft Office\Office15\outlook.exe"
        ;~ Return
    ;~ }

;-----------------------------------------------------------------------------
; Win+Shift+C = Open / Switch to Chrome
#+c::
    if WinExist("Google Chrome")
    {
        WinActivate
        return
    }
    else
    {
        Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        Return
    }
;-----------------------------------------------------------------------------
; Win+Shift+F = Open / Switch to Firefox
#+f::
    if WinExist("Mozilla Firefox")
    {
        WinActivate
        return
    }
    else
    {
        Run firefox
        Return
    }
;-----------------------------------------------------------------------------
; Win+Shift+A = Open / Switch to AutoCAD
#+a::
    if WinExist("ahk_class AfxMDIFrame140u")
    {
            WinActivate
            return
    }
    else
    {
        Run "C:\Program Files\Autodesk\AutoCAD 2020\acad.exe"
        Return
    }
;-----------------------------------------------------------------------------
; Win+C = Go to C Drive
#c::
    Run "C:\"
    Return
;-----------------------------------------------------------------------------
; Win+Shift+D = Open Documents 
#+D::
Run "C:\Users\Dylan\Documents"
return

;-----------------------------------------------------------------------------
; Win+Shift+M = Open Machine Learning Folder 
#+m::
Run "C:\Users\Dylan\Documents\Grad School stuff\Fall 2019\Machine Learning"
return

;-----------------------------------------------------------------------------
; Win+Shift+R = Open Repositories Folder 
#+r::
Run "C:\Users\Dylan\Documents\repos"
return

;-----------------------------------------------------------------------------
;Win+X = kill this script
#x::ExitApp
    