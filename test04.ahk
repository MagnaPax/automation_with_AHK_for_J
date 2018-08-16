#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include GetActiveBrowserURL.ahk
#Include FG_Update.ahk



global CurrentOrderIdNumber, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, CustomerNoteOnWebVal, StaffOnlyNoteVal, CompanyName, PendingOrderStatus, AlreadyProcessedItem, CurrentPONumber, text

; 이거 원래 GetActiveBrowserURL.ahk 파일 안에 있던 함수인데 이게 메인에 선언되야 com 으로 처리한 변수들의 값이 유지되어 메인에서 사용할 수 있다.
Menu, Tray, Icon, % A_WinDir "\system32\netshell.dll", 86 ; Shows a world icon in the system tray

ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass"


;~ Paymentwb.document.getElementsByTagName("INPUT")[3].innerText := Password  ;Password 입력
;~ Paymentwb.document.getElementsByTagName("INPUT")[6].Click() ; 로그인 버튼 누르기



FG_Update(arr)






wb := IEGet("My Drive - Google Drive - Internet Explorer")

; everything is not working
/*
wb.document.getElementsByTagName("SPAN")[84].focus()
Send, {AppsKey}            
Send, {Down}
Send, {Down}
Send, {Enter}
*/

;~ wb.document.getElementsByTagName("SPAN")[84].click(Right)
;~ wb.document.getElementsByTagName("SPAN")[84].rightclick()

MsgBox

            
            
            
			;~ wb.document.getElementsByTagName("A")[193].sendKeys(wb.Keys.ENTER)
            ;~ driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.ENTER)
			Send, {AppsKey}
            Send, {Down}
            Sleep 1000
            Send, {Down}
            Sleep 1000
            Send, {Down}
            Sleep 1000
            
            
            MsgBox

/*
;~ OpenCreateInvoiceTab()
;~ OpenCreateSalesOrdersSmallTab()

CurrentPONumber = MTR1CF3EC217A ;뉴오더에 있는 것
;~ CurrentPONumber = MTR2E511F454 ;뉴오더에 없는 것

OpenFGforNewFGProcessing_UsingChromeBySelenium()
MsgBox

;~ OpenFGforNewFGProcessing_NEW02()

;~ WinClose, ahk_class IEFrame
MsgBox, OUT

MsgBox, %CustomerNoteOnWebVal%`n`n%StaffOnlyNoteVal%
*/

Exitapp


Esc::
 Exitapp
 
 
 ^q::
SendInput, %text%
return
