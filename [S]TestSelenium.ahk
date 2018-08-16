
global CurrentOrderIdNumber, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, CustomerNoteOnWebVal, StaffOnlyNoteVal, CompanyName, PendingOrderStatus, AlreadyProcessedItem, CurrentPONumber, text, driver, sURL


;~ CurrentPONumber = MTR1D37B7E217
CurrentPONumber = MTR1D5A398C41
;~ CurrentPONumber = 12345678


;~ loginFG() ; FG 로그인 함수


OpenFGforNewFGProcessing_ChromeUsingSelenium()


;~ RE_OpenNewOrderPageToCheck()	; 함수 호출이 끝나면 크롬 창도 같이 닫히니까 인보이스 인쇄 전 한번 더 띄워서 확인하기 위해





	
OpenFGforNewFGProcessing_ChromeUsingSelenium(){
	
	;~ ###############################################################################################################
	;~ 처음에 Selenium으로 로그인 창을 열어서 로그인 한 뒤 프로필 파일이 저장된 경로를 얻으면 경로 끝이 \Default 로 끝난다.
	;~ 그 후 프로파일 경로를 사용키 위해 driver.SetProfile("프로파일경로") 를 실행하면 이상하게도 프로파일을 읽지 못한다
	;~ 또 다시 로그인을 해서 경로 끝을 \Default\Default 이렇게 \Default를 두 번 되게 만든 뒤
	;~ 다음 호출 시 \Default 한개로만 끝나는 경로를 읽어야만 그제서야 프로파일을 읽고 제대로 작동한다. 왜그런지는 모르겠음
	;~ ###############################################################################################################


	; PO 번호를 입력해서 검색 후 나오는 값을 얻기 위해 아래 코드를 쓴다.
	; if (driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/table/tbody/tr/td[5]/a").Attribute("outerHTML"))
	; 그런데 이 코드는 검색된 검색 결과 값이 없으면 에러가 난다. 이 에러를 무시하고 계속하면 그 다음 코드인 else가 실행된다. else에서 값이 없으면 이미 처리된 주문이라는 안내 띄워고 함수를 빠져나오는 처리를 하게 해놨다.
	; 에러 없는 제대로 된 코드를 어떻게 짜는 지 몰라서 아예 에러 경고창이 나오는 것을 없앴다. 에러가 발생해도 경고창에서 Yes 누르면 그 다음 코드 정상적으로 실행하기 때문에
;	ComObjError(false)
	

	; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
	BlockInput, Mouse
	
	Clipboard :=	
	ChromeProfile :=
	ValOFOuterHTML :=
	sURL :=	
	

	Loginname = customer3
	Password = Jo123456789
	URL = https://vendoradmin.fashiongo.net/#/order/orders/new ; 뉴오더 검색 페이지 열기
	

	; ChromeProfile.txt 내용을 읽어와서 ChromeProfile 변수에 저장하기
	FileRead, ChromeProfile, %A_ScriptDir%\CreatedFiles\ChromeProfile.txt


	driver := ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	
	;~ driver.SetProfile("C:\Users\Hahn\AppData\Local\Temp\Selenium\scoped_dir3536_5116\Default") ; 로그인을 또 하지 않기 위해서 크롬 프로파일 읽기
	driver.SetProfile(ChromeProfile) ; 로그인을 또 하지 않기 위해서 크롬 프로파일 읽기
	
	driver.Get(URL) ; 창 열기
	

	; 현재 url 얻는 부분
	sURL := driver.Url
	

	; 얻은 현재 url이 로그인 화면이면 로그인 하기
	if(RegExMatch(sURL, "imU)login")){		
		
		driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[1]/input").SendKeys(Loginname)	; Username
		driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[2]/input").SendKeys(Password)	; Password
		driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[4]/button").click()	; Button
		

		; 프로파일 정보(저장된 주소) 얻기
		Sleep 200
		Send, !d
		Sleep 200
		Send, chrome://version/
		Sleep 200
		Send, {Enter}
		
		; 프로파일 경로를 읽어서 ChromeProfile 변수에 저장한다
		ChromeProfile := driver.findElementByID("profile_path").Attribute("innerText")
		


		; 변수 안에 있는 \Default 중 한 개가 공백으로 교체된다
		; 이걸 해줘야 파일에 \Default 가 한 개만 있는 경로가 저장되고 다음에 driver.SetProfile(ChromeProfile) 실행할 때 제대로 읽게 된다.
		StringReplace, ChromeProfile, ChromeProfile, `\Default

		

		; 기존의 파일 내용을 초기화 하기 위해 ChromeProfile.txt 파일을 EmptyFile.txt 로 덮어씌우기
		FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\ChromeProfile.txt, 1
		
		Sleep 500

		; ChromeProfile 안에 있는 프로파일 주소 ChromeProfile.txt 파일에 쓰기(저장하기)
		FileAppend, %ChromeProfile%, %A_ScriptDir%\CreatedFiles\ChromeProfile.txt
		
		
		;~ MsgBox, % ChromeProfile
		
		; 안 닫아주면 아래에서 다시 열 때 오류 생기고 작동 안 함
		driver.quit()
		
		; 로그인을 한 뒤 함수를 재귀호출해서 다시 시작하기
		; 이렇게 안 하면 chrome://version/ 로 연 화면을 다시 뉴 오더 화면으로 돌려야 되는것 만들어야 되는데 그게 더 복잡할 듯
		OpenFGforNewFGProcessing_ChromeUsingSelenium()

	
	}

;~ driver.executeScript("driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS)") ; 이것도 위와 똑같은 작동 하는데 자바 스크립트를 직접적으로 주입(?)한 코드 표현법
;~ driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS)
	; 드롭다운 메뉴 PO Number 로 바꾸기	
	;~ driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div[1]/div[2]/select").click() ; 드롭다운 메뉴 선택(클릭)하기
	;~ driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div[1]/div[2]/select").SendKeys(driver.Keys.ArrowDown) ; 아래로 내리기
	;~ driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div[1]/div[2]/select").sendKeys(driver.Keys.TAB) ; 탭키 누르기
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div[1]/div[2]/select").SendKeys("PO Number")

	; 검색란에 PO 번호 입력
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div[1]/div[3]/input").SendKeys(CurrentPONumber)

	; Apply 버튼 클릭
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div[1]/button").click()

	Sleep 300 ; 값을 읽는데 시간이 걸리는 것 같음. 이거 없으면 오류남


	; 검색된 PO Number 값이 있는지 확인하고
	if (driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/table/tbody/tr/td[5]/a").Attribute("outerHTML"))
	{
		; 검색된 값이 있으면 클릭해서 들어가기
		ValOFOuterHTML := driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/div[3]/div/fg-order-list/div[2]/table/tbody/tr/td[5]/a").click()
		
		ReadingAndProcessingOnFGForNewOrdersUsingSelenium() ; 클릭해서 들어간 페이지에서 처리하기 위해
		
		
		return ; 이미 위 함수에서 처리했기 때문에 밑으로 진행할 필요 없이 그냥 함수 끝내고 리턴
	}	
	else ; 검색된 값이 없으면 Ship Today 등으로 이미 처리된 PO 번호라는 뜻. ; 혹시 모르니 확인 차원에서 LAMBS 창 띄워주고 함수 빠져나가기
	{
		
		; PO 번호 없을 때 LAMBS 열어서 확인하는 함수 호출
		CheckLAMBSWhenNoPOnumberResult()

		; 이미 처리된 주문이라는 표시 해주기 위해 AlreadyProcessedItem 값을 1로 만듬
		AlreadyProcessedItem = 1
		
		return
	}

	
	MsgBox, % ValOFOuterHTML

	
	return
}	


	
; PO 번호 없을 때 LAMBS 열어서 확인하는 함수
CheckLAMBSWhenNoPOnumberResult(){
	
	; Create Sales Orders Small 열기
;	OpenCreateSalesOrdersSmallTab()
		
	;Hide All 클릭해서 메뉴 바 없애기
;	ClickAtThePoint(213, 65)

	;Account Summary 열기
;	MouseClick, l, 920, 155
;	WinWaitActive, Accounts Summary
;	WinMaximize
		
	; BO 목록표 열기 (Customer Order + 버튼 클릭)
;	ControlClick, WindowsForms10.BUTTON.app.0.378734a4, Accounts Summary
;	WinWaitActive, Customer Order +Zoom In
;	WinMaximize
	
	;~ MsgBox, 262144, Already Processed Order, THIS ORDER ALREADY HAS BEEN PROCESSED.`n`n`nTHIS WINDOW WILL BE CLOSED IN 3 SECONDS, 3
	MsgBox, 262144, Already Processed Order, The Number is %CurrentOrderIdNumber%`nTHIS PO NUMBER MIGHT BE ALREADY PROCESSED.
		
	WinClose, Customer Order +Zoom In
	WinClose, Accounts Summary
	WinClose, FashionGo Vendor Admin - Internet Explorer
			
	
	return
}



; PO 번호로 열린 뉴오더 페이지에서 Status 바꾸고 정보 읽는 등 여러가지 처리하기
ReadingAndProcessingOnFGForNewOrdersUsingSelenium(){

	
	Sleep 400
	; Buyer Notes 읽기
	CustomerNoteOnWebVal := driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[1]").Attribute("value")
	StringUpper, CustomerNoteOnWebVal, CustomerNoteOnWebVal ; 고객 메모 대문자로 바꾸기	
	
	
	Sleep 400
	; Staff Notes (Internal Use Only) 읽기
	StaffOnlyNoteVal := driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[2]/div/textarea").Attribute("value")
	StringUpper, StaffOnlyNoteVal, StaffOnlyNoteVal ; Staff only notes 대문자로 바꾸기


	Sleep 400
	; Additional Info 읽기
	AdditionalInfo := driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[2]").Attribute("value")
	StringUpper, AdditionalInfo, AdditionalInfo ; 대문자로 바꾸기



	Sleep 300 ; 값을 읽는데 시간이 걸리는 것 같음. 이거 없으면 오류남	
	; Order Status 를 Confirmed Orders 로 바꾸기
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/span/select").SendKeys("Confirmed Orders")
	
	
	Sleep 300
	; Update & Notify Buyer 버튼 클릭
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div[1]/ul/li[2]/span/button[2]").click()
	
	
	Sleep 300
	; Send Mail 버튼 클릭
;	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/fg-notify-buyer-modal/div/div[2]/div/div/div[2]/fieldset/div[5]/button[2]").click()
	


	;~ driver.FindElementByXPath("").click()
	;~ driver.FindElementByXPath("").SendKeys(driver.Keys.ArrowDown)
	;~ driver.FindElementByXPath("").sendKeys(driver.Keys.TAB)
	;~ driver.FindElementByXPath("").Attribute("innerText")


	MsgBox, CustomerNoteOnWebVal : %CustomerNoteOnWebVal%`n`nStaffOnlyNoteVal : %StaffOnlyNoteVal%
	
	; 현재 창 url 읽어서 전역변수 sURL에 저장
	sURL := driver.Url


	;~ driver.open()
	;~ Thread.sleep()
	
	return
}




; 함수 호출이 끝나면 크롬 창도 같이 닫히니까 인보이스 인쇄 전 한번 더 띄워서 확인하기 위해
RE_OpenNewOrderPageToCheck(){
	
	
	; ChromeProfile.txt 내용을 읽어와서 ChromeProfile 변수에 저장하기
	FileRead, ChromeProfile, %A_ScriptDir%\CreatedFiles\ChromeProfile.txt

	
	
	driver := ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	
	;~ driver.SetProfile("C:\Users\Hahn\AppData\Local\Temp\Selenium\scoped_dir3536_5116\Default") ; 로그인을 또 하지 않기 위해서 크롬 프로파일 읽기
	driver.SetProfile(ChromeProfile) ; 로그인을 또 하지 않기 위해서 크롬 프로파일 읽기
	
	driver.Get(sURL) ; 창 열기
	
	
	; Additionla Info 칸을 아래로 내리기 (주문 내용을 확인하기 위해 화면을 아래로 내리는 효과)
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[2]").click()
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[2]").SendKeys(driver.Keys.ArrowDown)
	
	
	
	MsgBox, re open
	
	return
}
























loginFG(){
	
	
	Loginname = customer3
	Password = Jo123456789
	DisableInfobars = disable-infobars
	URL = https://vendoradmin.fashiongo.net/#/auth/login

	
	
	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	;~ driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	driver.AddArgument(DisableInfobars) ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	driver.Get(URL) ; 창 열기
	

	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[1]/input").SendKeys(Loginname)	; Username
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[2]/input").SendKeys(Password)	; Password
	driver.FindElementByXPath("/html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[4]/button").click()	; Button

	
	; 프로파일 정보(저장된 주소) 얻기
	Sleep 200
	Send, !d
	Sleep 200
	Send, chrome://version/
	Sleep 200
	Send, {Enter}
	
	; 프로파일 경로를 읽어서 ChromeProfile 변수에 저장한다
	ChromeProfile := driver.findElementByID("profile_path").Attribute("innerText")
	

	; 기존의 파일 내용을 초기화 하기 위해 ChromeProfile.txt 파일을 EmptyFile.txt 로 덮어씌우기
	FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\ChromeProfile.txt, 1
		
	Sleep 500

	; ChromeProfile 안에 있는 프로파일 주소 ChromeProfile.txt 파일에 쓰기(저장하기)
	FileAppend, %ChromeProfile%, %A_ScriptDir%\CreatedFiles\ChromeProfile.txt
		
	driver.quit()
	
	
	;~ MsgBox, % ChromeProfile
	
	return
}



	
	
	


	Exitapp

	Esc::
	 Exitapp	