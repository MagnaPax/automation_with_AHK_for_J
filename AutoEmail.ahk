; #################################################################
; 기본적으로 Auto Email, Divison, Warehouse 업데이트 하고 
; 메모가 없으면 Ship via, Terms, Pay.Method, Priority 업데이트 하기
; #################################################################


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



#Include %A_ScriptDir%\lib\

#Include N41.ahk
#Include FindTextFunctionONLY.ahk

N_driver := new N41


MsgBox, 262144, Title, Ok to Start

Sleep 9000



; Auto Email, Divison, Warehouse 그리고 Ship via, Terms, Pay.Method, Priority 업데이트 하기
N_driver.UpdateInfoForAutoEmail()




ExitApp





;~ Esc::
;~ Space::
Esc::
ExitApp

F5::
Reload