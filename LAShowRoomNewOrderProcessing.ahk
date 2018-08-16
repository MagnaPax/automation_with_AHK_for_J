#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.















^e::
CompanyName = %Clipboard%
StringUpper, CompanyName, CompanyName ; Staff only notes 대문자로 바꾸기
SendInput, %CompanyName%
return

GuiClose:
 ExitApp

Esc::
 Xl.ActiveWorkbook.save()
 Exitapp
 