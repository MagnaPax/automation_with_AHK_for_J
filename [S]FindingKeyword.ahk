#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



WinClose, Pick Ticket Print
;~ SoundPlay, %A_WinDir%\Media\Ring04.wav
;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
SoundPlay, %A_WinDir%\Media\Ring07.wav
;~ SoundPlay, %A_WinDir%\Media\Alarm01.wav
MsgBox

/*
driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
driver.Get("http://the-automator.com")

;********************find and click text***********************************
keyword:="Continue reading"  ; keyword you want to find- Note Case sensitive

if(driver.FindElementByXPath("//*[text() = '" keyword "']"))
	driver.FindElementByXPath("//*[text() = '" keyword "']").click()

MsgBox, pause
*/



;~ /*

URL = http://the-automator.com
URL = https://vendoradmin.fashiongo.net/#/order/orders
run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)




; 열려있는 창 재사용 하기 (반드시 코멘드 라인이 --remote-debugging-port=9222 로 열린 창만 재사용 가능
driver := ChromeGet()
MsgBox, % driver.Window.Title "`n" driver.Url



keyword = Confirmed

if(driver.FindElementByXPath("//*[text() = '" keyword "']"))
	driver.FindElementByXPath("//*[text() = '" keyword "']").click()


*/



Exitapp

Esc::
 Exitapp




ChromeGet(IP_Port := "127.0.0.1:9222") {
	driver := ComObjCreate("Selenium.ChromeDriver")
	driver.AddArgument("start-maximized") ; 윈도우 최대창으로 만들기
	driver.SetCapability("debuggerAddress", IP_Port)
	driver.Start()
	return driver
}



