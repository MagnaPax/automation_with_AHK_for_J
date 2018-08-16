#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include %A_ScriptDir%\lib\

#Include function.ahk


; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
BlockInput, Mouse


Run, C:\COMP-SYS\DLL\LAMBS.exe


WinWaitActive, Login
ControlSetText, WindowsForms10.EDIT.app.0.378734a3, CHUNHEE, Login
ControlSetText, WindowsForms10.EDIT.app.0.378734a2, 5425, Login
ControlClick WindowsForms10.BUTTON.app.0.378734a2, Login, , l

WinWaitClose, Status


WinWaitActive, LAMBS -  Garment Manufacturer & Wholesale Software
WinWaitClose, Status

while (A_cursor = "Wait")
	Sleep 3000

OpenStyleMasterTab()
WinWaitClose, Status


; 커서 상태가 작업처리중이면 끝날때까지 기다리기
while (A_cursor = "Wait")
	Sleep 3000
;~ Sleep 2000


OpenCustomerInfoTab()
WinWaitClose, Status


; 커서 상태가 작업처리중이면 끝날때까지 기다리기
while (A_cursor = "Wait")
	Sleep 3000
;~ Sleep 2000


OpenCreateSalesOrdersSmallTab()
WinWaitClose, Status


; 커서 상태가 작업처리중이면 끝날때까지 기다리기
while (A_cursor = "Wait")
	Sleep 3000
;~ Sleep 2000


OpenCreateInvoiceTab()
WinWaitClose, Status


; 커서 상태가 작업처리중이면 끝날때까지 기다리기
while (A_cursor = "Wait")
	Sleep 3000




MsgBox, , , IT'S DONE`n`n`nTHIS WINDOW WILL BE CLOSED IN 3 SECONDS, 3






Exitapp

Esc::
 Exitapp
 Reload
