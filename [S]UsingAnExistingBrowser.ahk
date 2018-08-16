#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

global driver

; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)


;~ Using_ELAMBS_To_Find_An_Item()
;~ MsgBox

GetCCInfo()
MsgBox

GetCCInfo(){
	
	driver := ChromeGet()
	
	;~ MsgBox, % driver.Window.Title "`n" driver.Url
	
	
	; 메뉴바에 있는 Accounts 클릭
	Xpath = //*[@id="topmenu-header"]/li[1]/a
	driver.FindElementByXPath(Xpath).click()
	

	; 위에서 메뉴바의 Accounts 를 클릭 후 하위 메뉴인 Customer List 항목이 나타날 때까지 대기하다가 나타나면 클릭
	Xpath = //*[@id="topmenu-body"]/ul[1]/li[1]/ul/li[1]/a
	Wait_Until_Element_Is_Visible(Xpath) ; element 가 나타날 때까지 대기
	driver.FindElementByXPath(Xpath).click()
	
	
	; Search 버튼 클릭 
	Xpath = //*[@id="contents-header"]/div[4]
	driver.FindElementByXPath(Xpath).click()
	
	
	; 위에서 Search 버튼을 클릭한 후 상세 메뉴인 Search Condition 입력창이 나타날 때까지 대기하다가 나타나면 그곳에 값 입력
	; 값 입력 후 엔터
	Xpath = //*[@id="ContentPlaceHolder1_SearchQuery"]
	PutInString = THE MINT JULEP BOUTIQUE LLC
	driver.FindElementByXPath(Xpath).SendKeys(PutInString)
	driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.ENTER)
		

	; 값 입력 후 검색 결과가 나타날 때까지 기다리다가 검색된 결과가 나타나면 우클릭
	Xpath = //*[@id="contents-body"]/div/div[2]/table/tbody/tr[3]/td[2]/a
	Wait_Until_Element_Is_Visible(Xpath) ; element 가 나타날 때까지 대기
	
	driver.FindElementByXPath(Xpath).rightclick()
	;~ Click Right
	;~ Send {Click , , right}  ; Control+RightClick
	;~ driver.FindElementByXPath(Xpath).contextClick().build().perform()
	;~ parentMenu  := driver.FindElementByXPath(Xpath)
	
	;~ driver.contextClick(parentMenu).build().perform() ; //Context Click
	;~ driver.contextClick(parentMenu).context_click(element).perform()
	
	;~ driver.FindElementByXPath(Xpath).contextClick(Xpath).build().perform()

	Sleep 1000
	driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.RButton)
	Sleep 1000
	driver.FindElementByXPath(Xpath).SendKeys(driver.Keys.ArrowDown).SendKeys(driver.Keys.ArrowDown)
	Sleep 1000
	driver.FindElementByXPath(Xpath).SendKeys(driver.Keys.ArrowDown)
	;~ driver.FindElementByXPath(Xpath).SendKeys(driver.Keys.ArrowDown)
	;~ driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.TAB)


	
	
	return
}

Using_ELAMBS_To_Find_An_Item(){
	
	URL = https://www.freecrm.com/index.html
	
	;~ run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)
	
	driver := ChromeGet()
	
	MsgBox, % driver.Window.Title "`n" driver.Url
	

	;~ Xpath = //a[text()='Features']
	;~ Xpath = //*[@id="navbar-collapse"]/ul/li[1]/a
	
	; 메뉴바에 있는 Items 클릭
	Xpath = //*[@id="topmenu-header"]/li[2]/a	
	driver.FindElementByXPath(Xpath).click()


	; 위에서 메뉴바의 Items 을 클릭 후 하위 메뉴인 Item List 항목이 나타날 때까지 대기하다가 나타나면 클릭
	Xpath = //*[@id="topmenu-body"]/ul[2]/li[1]/ul/li[1]/a
	Wait_Until_Element_Is_Visible(Xpath) ; element 가 나타날 때까지 대기
	driver.FindElementByXPath(Xpath).click()
	
	
	; Search 버튼 클릭 
	Xpath = //*[@id="contents-header"]/div[5]
	driver.FindElementByXPath(Xpath).click()


	; 위에서 Search 버튼을 클릭한 후 상세 메뉴인 Search Condition 입력창이 나타날 때까지 대기하다가 나타나면 그곳에 값 입력
	; 값 입력 후 엔터
	Xpath = //*[@id="search-query"]
	PutInString = P2429
	PutInString = b1234
	driver.FindElementByXPath(Xpath).SendKeys(PutInString)
	driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.ENTER)
	

	; 값 입력 후 검색 결과가 나타날 때까지 기다리다가 검색된 결과가 나타나면 클릭
	Xpath = //*[@id="contents-body"]/div/div[2]/table/tbody/tr[3]/td[2]/a
	Wait_Until_Element_Is_Visible(Xpath) ; element 가 나타날 때까지 대기
	
	;~ if (driver.FindElementByXPath(Xpath).Attribute("innerText") == "")
	if (!driver.FindElementByXPath(Xpath).Attribute("innerText"))
	{
		MsgBox, Not found
		;~ return	
	}
	
	driver.FindElementByXPath(Xpath).click()

	
	return
}


T2(){
	
	URL = http://the-automator.com/
	
	;~ run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)
	
	driver := ChromeGet()
	
	MsgBox, % driver.Window.Title "`n" driver.Url


	Xpath = //*[@id="menu-item-1662"]/a/span
	
	driver.FindElementByXPath(Xpath).click()	

	
	return
}




Test_Xpath()
;~ MsgBox, pause


UsingAnExistingBrowser()
MsgBox, function out


; 열려있는 창 재사용 하기 (반드시 코멘드 라인이 --remote-debugging-port=9222 로 열린 창만 재사용 가능
driver := ChromeGet()
MsgBox, % driver.Window.Title "`n" driver.Url


; 이 함수로 열린 크롬 창은 함수가 종료되면 없어짐
Open()
MsgBox, Pause


Exitapp




Test_Xpath(){

	;~ URL = http://naver.com
	;~ URL = https://www.freecrm.com/index.html
	URL = http://the-automator.com/


	run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)

	; 열려있는 창 재사용 하기 위해 함수 호출
	driver := ChromeGet()

	driver.findElementsByName("s").item[1].SendKeys("hello world")
	driver.findElementsByName("s").item[1].SendKeys(driver.Keys.ENTER) ;http://seleniumhome.blogspot.com/2013/07/how-to-press-keyboard-in-selenium.html
	MsgBox pause
	
	
	;~ driver.FindElementByXPath("").click()
	MsgBox, click pause


	return
}


UsingAnExistingBrowser(){	

	;~ URL = http://naver.com
	URL = http://the-automator.com/


	;~ ############## --new-window 이 command line 은 새 창으로 열기 ##############
	;~ ############## --remote-debugging-port=9222 이 command line 은 ChromeGet() 함수 이용하면 이 코멘드 라인으로 열린 창을 이용할 수 있게 됨 ##############

	;~ run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL ;일반적인 크롬 새 창으로 열기
	;~ run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --remote-debugging-port=9222 " : " " ) URL   ;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (기존에 열린 창이 있으면 새 탭으로 열림)
	run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)

	; 열려있는 창 재사용 하기 위해 함수 호출
	driver := ChromeGet()

	driver.findElementsByName("s").item[1].SendKeys("hello world")	
	driver.findElementsByName("s").item[1].SendKeys(driver.Keys.ENTER) ;http://seleniumhome.blogspot.com/2013/07/how-to-press-keyboard-in-selenium.html
	MsgBox pause

	return
}


Open(){
driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
driver.Get("http://the-automator.com/")
MsgBox, % driver.Window.Title "`n" driver.Url
return
}



ChromeGet(IP_Port := "127.0.0.1:9222") {
	driver := ComObjCreate("Selenium.ChromeDriver")
	driver.SetCapability("debuggerAddress", IP_Port)
	driver.Start()
	return driver
}


Wait_Until_Element_Is_Visible(Xpath){
;~ /*	
	;~ MsgBox, % Xpath
	
	NoProductsFound = //*[@id="contents-body"]/div/div[2]/div

	loaded := false
	While !loaded
	{
		try
		{
			
	;~ if (A_TickCount - StartTime > 2*MaxTime + 100)
	;~ {
		;~ MsgBox 너무 많은 시간지 경과하였습니다.
		;~ ExitApp
		;~ return
	;~ }		
			
			;~ if (driver.FindElementByXPath(NoProductsFound).Attribute("innerText") = "No Products Found"){
				;~ MsgBox, Not found
				;~ return
			;~ }			
			
			;~ if (ie.document.getElementById("some_ID").innertext != "") ; 이건 IE 열어서 Com Object 사용할 때
			if (driver.FindElementByXPath(Xpath).Attribute("innerText") != "")
				loaded := true
			

		}
		Sleep 500
	}
*/

NoProductsFound = //*[@id="contents-body"]/div/div[2]/div

; 나타나길 기다리는 element 값이 거짓 '!' 일 동안 계속 sleep 500 반복
; 즉, 나타날 때까지 기다리기
;~ while(!driver.FindElementByXPath(Xpath).Attribute("innerText"))
	;~ Sleep 500

;~ {
	;~ Sleep 500
	
	;~ if (A_TickCount - StartTime > 2*MaxTime + 100)
;~ {
    ;~ MsgBox 너무 많은 시간지 경과하였습니다.
    ;~ ExitApp
	;~ return
;~ }
	
	
	;~ if (driver.FindElementByXPath(NoProductsFound).Attribute("innerText") = "No Products Found"){
		;~ MsgBox, Not found
		;~ return
	;~ }			
;~ }




	return
}




Esc::
;~ WinClose, ahk_exe chrome.exe
Exitapp