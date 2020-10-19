; -*- coding: utf-8 -*-
; File: eid_viewer_export
; Author: Theo Technicguy eid_viewer_export-program@licolas.net
; Interpreter: AutoHotkey
; Ext: shk
; Licenced under GPU GLP v3. See LICENCE file for information.
; Copyright (c) TheoTechnicguy 2020.
; Version: 0.1.0
; -----------------------
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

if WinExist("ahk_exe eID Viewer.exe"){
  WinActivate
} else {
  Run, "C:\Program Files (x86)\Belgium Identity Card\EidViewer\eID Viewer.exe"
}

; -------------
; Wait loading

; From AHK Help manual
; CoordMode Pixel  ; Interprets the coordinates below as relative to the screen rather than the active window.
; ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *Icon3 %A_ProgramFiles%\SomeApp\SomeApp.exe ; THIS need to change
; if (ErrorLevel = 2)
;     MsgBox Could not conduct the search.
; else if (ErrorLevel = 1)
;     MsgBox Icon could not be found on the screen.
; else
;     MsgBox The icon was found at %FoundX%x%FoundY%.


Send, ^s

WinActivate, "ahk_class #32770"
WinWaitActive
; WARNING: Ensure %TMP%/eid exists!
; Sleep, 5000
Send, % "%TMP%\eid\patient.eid" ; Check save
Send, {Enter}
