#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


/*
Java 에서는 아래처럼 구현한다
https://stackoverflow.com/questions/11428026/select-an-option-from-the-right-click-menu-in-selenium-webdriver-java
*/



/*
URL = http://jodifl.elambs.com/page_Sales/Invoice_list.aspx
run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)
MsgBox
*/




driver := ChromeGet()
MsgBox, % driver.Window.Title "`n" driver.Url

;~ Xpath = //*[@id=":75.0B3q2fVZ-KzYFMDU3Y1VvSDJ0RFk"]/div[1]/div/div[2]/div/div[3]/span
Xpath = //*[@id="contents-body"]/div/div[2]/table/tbody/tr[11]/td[2]/a


; 아래는 자바 코드(ahk와 비교하기 위해서)
;~ Actions action= new Actions(driver);
;~ action.contextClick(productLink).build().perform();


element := driver.FindElementByXPath(Xpath)
driver.Actions.ClickContext(element).Perform()
Sleep 1000




/* eLAMBS 에서는 메뉴의 각 항목마다 고유의 Xpath 주소가 있어서 이렇게 직접 Xpath를 이용해서 클릭해주면 됨
Xpath = /html/body/div[2]/div[11]/a
driver.FindElementByXPath(Xpath).click()
MsgBox
*/


/* 구글 드라이브에서는 메뉴의 항목들이 고유한 Xpath가 없고 같은 Xpath를 공유하고 있어서 아래처럼 화살표를 한칸씩 이동한 뒤 선택해줘야 됨
driver.Actions.SendKeys(driver.Keys.ArrowDown).perform()
driver.Actions.SendKeys(driver.Keys.Enter).perform()
*/






; 이미 존재하는 크롬 창 제어하기 위한 함수. 아래와 같은 코멘드 라인으로 열린 크롬이라야 이 함수로 컨트롤 가능하다
; Start Chrome with command line: chrome.exe --remote-debugging-port=9222
ChromeGet(IP_Port := "127.0.0.1:9222") {
	driver := ComObjCreate("Selenium.ChromeDriver")
	driver.SetCapability("debuggerAddress", IP_Port)
	driver.Start()
	return driver
}





Exitapp

Esc::
Exitapp


