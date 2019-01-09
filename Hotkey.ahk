;Author: Dylan Rosser 2018
;Cosentini Associates

; This is a script that will run nor switch to certain applications. It will also launch common directories like the C drive or P Dump.
;Add as many of these as you want to a .ahk file, and set that to be run at startup.

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
 #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
#SingleInstance
;SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; See the Hotkeys reference [1] for details of the modifiers and keys available.
; [1]: http://www.autohotkey.com/docs/Hotkeys.html

;   # - Windows Key
;   + - Shift Key
;   ^ - Control Key
;   ! - Alt Key



;---------------------------------------------------------------------------------------------------------------------------------------------------
;Win+Shift+q = tells the class of the active window (used for debugging)
+#q::
WinGetClass, class, A
MsgBox, The active window's class is "%class%".
return
;---------------------------------------------------------------------------------------------------------------------------------------------------
;Win+Shift+D = Remind Everyone that Dylan is Awesome
+#D::
SetTimer, ChangeButtonNames, 50 
MsgBox, 52, Awesomeness Alert!, Dylan is Awesome
return 
ChangeButtonNames: 
IfWinNotExist, Awesomeness Alert!
    return  ; Keep waiting.
SetTimer, ChangeButtonNames, Off 
WinActivate 
ControlSetText, Button1, &I Agree 
ControlSetText, Button2, &Definately 
return
;--------------------------------------------------------------------------------------------------------------------------------------------------
;Win+M = Minimize the active Window
#m::WinMinimize, A
return
;--------------------------------------------------------------------------------------------------------------------------------------------------
;CTRL+Q = Quit the Active Application
^q::WinClose, A
return
;--------------------------------------------------------------------------------------------------------------------------------------------------
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
;--------------------------------------------------------------------------------------------------------------------------------------------------
;Win+V = Run spreadsheet with VBA functions for Electrical Engineering Calculations
#v::
if WinExist("MyMacros")
    {
        WinActivate
        return
    }
else
    {
        Run "C:\Users\dylan.rosser\Desktop\References\VBA Macro's\MyMacros.xlsm"
        return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
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
;---------------------------------------------------------------------------------------------------------------------------------------------------
;Win+Shift+X = Open an active excel sheet or launch excel
+#x::
if WinExist("ahk_class XLMAIN")
    {
        WinActivate
        return
    }
else
    {
        Run "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.exe"
        return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
;Win+Shift+W= Open an active word file or launch MS Word
+#w::
if WinExist("ahk_class OpusApp")
    {
        WinActivate
        return
    }
else
    {
        Run "C:\Program Files (x86)\Microsoft Office\Office15\WinWord.exe"
        return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
;Alt+C = Open / Switch to calculator
!c::
if WinExist("Calculator")
    {
        WinActivate
        return
    }
    else 
    {
        Run "C:\Windows\System32\calc.exe"
        return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+Shift+r - Open Revit 2017
#+r::
    if WinExist("Autodesk Revit")
    {
            WinActivate
            ;WinMaximize
            return
    
    }
    else
    {
        Run "C:\Program Files\Autodesk\Revit 2017\Revit.exe"
        Return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+Shift+8 - Open Revit 2018
#+8::
    if WinExist("Autodesk Revit")
    {
            WinActivate
            ;WinMaximize
            return
    
    }
    else
    {
        Run "C:\Program Files\Autodesk\Revit 2018\Revit.exe"
        Return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+Shift+a - Open/Switch to AutoCAD 2015
#+a::
    if WinExist("Autodesk AutoCAD 2015")
    {
        WinActivate
        ;WinMaximize
        return
    }
    else
    {
        Run "C:\Program Files\Autodesk\AutoCAD 2015\acad.exe"
        Return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+Shift+O = Open Outlook
#+o::
    if WinExist("ahk_class rctrl_renwnd32")
    {
        WinActivate
        return
    }
    else 
    {
        Run "C:\Program Files (x86)\Microsoft Office\Office15\outlook.exe"
        Return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+Shift+B = Open / Switch to BlueBeam
#+b::
    if WinExist("ahk_class WindowsForms10.Window.3.app.0.73673b_r14_ad1")
        {
            WinActivate
            return
        }
        else if WinExist("ahk_class WindowsForms10.Window.3.app.0.73673b_r12_ad1")
        {
            WinActivate
            return
        }
        else
        {
            Run "C:\Program Files (x86)\Bluebeam Software\Bluebeam Revu\2017\Revu\Revu32.exe"
            Return
        }
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+Shift+C = Open / Switch to Chrome
#+c::
    if WinExist("ahk_class Chrome_WidgetWin_1")
    {
            WinActivate
            return
    }
    else
    {
        Run "C:\Users\dylan.rosser\AppData\Local\Google\Chrome\Application\chrome.exe"
        Return
    }
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+Shift+Break = Edit this file
#+Break::
    Run "C:\Program Files\AutoHotkey\SciTE\scite.exe" "C:\Users\dylan.rosser\Desktop\References\HotKeys\Hotkey.ahk"
    Return
;--------------------------------------------------------------------------------------------------------------------------------------------------    
; Win+C = Go to C Drive
#c::
    Run "C:\"
    Return
;--------------------------------------------------------------------------------------------------------------------------------------------------
; Win+I = Go to I Drive
#i::
    Run "I:\"
    Return
;--------------------------------------------------------------------------------------------------------------------------------------------------   
; Win+P = Go to P Dump
#p::
    Run "P:\DUMP"
    Return
;--------------------------------------------------------------------------------------------------------------------------------------------------
;Ctrl+Alt+R = Go to References Folder
^!r::
    Run "C:\Users\dylan.rosser\Desktop\References"
    Return
;--------------------------------------------------------------------------------------------------------------------------------------------------   
;Win+Shift+S = Go to paper space, zoom extents, save, and close an autocad file
#+s::
    SendInput,tilemode{enter}0{enter}zoom{enter}extents{enter}_qsave{enter}close{enter}{enter}
    Return
    
;--------------------------------------------------------------------------------------------------------------------------------------------------   
;Win+Shift+v = Automatically Navigate Pastespecial window to paste excel spreadsheets
#+v::
    SendInput,pastespec{enter}
        Send, !l
        Send, {enter}
        Send, {lbutton}
        Sleep, 500
        Send, scale{enter}
        Sleep, 100
        Send, {lbutton}
        Sleep, 100
        Send, {enter}
        Sleep, 500
        Send, {lbutton}
        Sleep, 100
        Send, 96{enter}
      
    Return
;--------------------------------------------------------------------------------------------------------------------------------------------------
;--------------------------------------------------------------------------------------------------------------------------------------------------
;Win+X = kill this script
#x::ExitApp
    