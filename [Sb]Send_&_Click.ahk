

	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	;~ driver:= ComObjCreate("Selenium.FireFoxDriver")
	driver.Get("http://the-automator.com/")
	
	
	PutInString = hello world2

	
	;~  !!!!DO NOT USE!!!  driver.findElement(By.id("search_form_input_homepage")).SendKeys("hello") ; Java bindings of Selenium. !!!!DO NOT USE!!!

;~ /*
	; item[1]은 s라는 이름이 두 개 있어서 첫 번째 칸에 입력하기 위해 값을 1로 설정
	;~ driver.findElementsByTagName("INPUT").arguments[0].SendKeys("hello world") ; 태그로 찾아가는 방법을 알아야 되는데
	driver.findElementsByName("s").item[1].SendKeys("hello world")	
	driver.findElementsByName("s").item[1].SendKeys(driver.Keys.ENTER) ;http://seleniumhome.blogspot.com/2013/07/how-to-press-keyboard-in-selenium.html
	MsgBox pause
	
	; 이건 두 번째 칸에 입력
	driver.findElementsByName("s").item[2].SendKeys("2nd")
	driver.findElementsByName("s").item[2].SendKeys(driver.Keys.ENTER) ;http://seleniumhome.blogspot.com/2013/07/how-to-press-keyboard-in-selenium.html
	MsgBox pause2
	
	; 이런 표현식도 가능
	driver.executeScript("arguments[0].setAttribute('value', 'hello world')", driver.findElementsByName("s")) ;sets value
	driver.executeScript("arguments[1].setAttribute('value', 'hello world2')", driver.findElementsByName("s")) ;sets value
	MsgBox pause2

	; 이렇게 findElementsByName 말고 Class 값으로도 위치 지정해서 값 입력 할 수 있음
	;~ driver.findElementsByClass("s").item[1].SendKeys("hello world") ; 바로 밑과 같은 동작
	driver.executeScript("arguments[0].setAttribute('value', 'hello world')", driver.findElementsByClass("s")) ; 이거나 바로 위나 모두 같은 곳에 값 입력
	driver.executeScript("arguments[1].setAttribute('value', 'hello world2')", driver.findElementsByClass("s")) ;두 번째 s 에 값 입력
	MsgBox pause2
*/

	; Xpath를 사용해서 접근할 수도 있음
	;~ driver.FindElementByXPath("//*[@id="search-6"]/form/label/input").SendKeys(username)
	;~ driver.FindElementByXPath("//*[@id=`"prime_nav`"]/li[10]/form/label/input").SendKeys("hello world")
	;~ driver.FindElementByXPath("//*[@id="prime_nav"]/li[10]/form/label/input").SendKeys("hello world")
	;~ driver.FindElementByXPath("//*[@id="prime_nav"]/li[10]/form/label/input").SendKeys(PutInString)
	;~ MsgBox pause2
	
	



	; 드롭다운 박스를 먼저 클릭 한 뒤
	driver.findElementByID("cat").click() ;1 based, not zero	
	;~ aa := driver.findElementByID("cat") ;이렇게 해도 위의 코드와 똑같은 동작을 한다
	;~ aa.click() ;이렇게 해도 위의 코드와 똑같은 동작을 한다
	
	
	
	;~ driver.findElementByID("cat").focus() ;1 based, not zero
	;~ driver.findElementByID("cat").contextClick(element).perform() ;1 based, not zero
	;~ MsgBox
	
	;~ MsgBox

	Loop, 7
	{
		; 칸을 한 개씩 내린다
		driver.findElementByID("cat").SendKeys(driver.Keys.ArrowDown)		
	}
	
	; 그 값을 선택하기 위해 tab 키 누름( 엔터 눌러도 됨)
	driver.findElementByID("cat").sendKeys(driver.Keys.TAB)


	; 선택한 드롭다운 박스 메뉴의 값을 읽어오기
	MsgBox % driver.findElementByID("cat").Attribute("value")



	Exitapp

	Esc::
	 Exitapp