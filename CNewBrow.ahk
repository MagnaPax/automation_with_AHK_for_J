



global URL


; ############################################################################################################################################################################################################################################################################
; ###########################################################################################################     넘겨받은 URL 로 이동하기     ################################################################################################################################
; ############################################################################################################################################################################################################################################################################


; url 받아서 이동하기. 필요하다면 로그인 하기
;~ goToURl_AfterLogIn_IfNeeded(driver, URL){
goToURl_AfterLogIn_IfNeeded(driver, URL){
	


	;~ driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	;~ driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	;~ driver.AddArgument("--start-maximized") ; 창 최대화 하기
	
		
	; chrome_Profile.txt 파일 내용을 읽어서 chrome_Profile 변수에 저장하기
	FileRead, chrome_Profile, %A_ScriptDir%\CreatedFiles\NewWindowProcessing\chrome_Profile.txt

;~ MsgBox, % "chrome_Profile`n`n" . chrome_Profile

	; 파일에서 읽어온 프로파일 경로를 읽어서 적용하기
	driver.SetProfile(chrome_Profile)

	; url로 이동해 보기	
	driver.Get(URL)



	; 만약 현재 페이지가 FG 페이지라면
	if(RegExMatch(driver.Url, "imU)fashiongo")){
		
		; FG 로그인 페이지라면
		if(RegExMatch(driver.Url, "imU)login")){
			
			; driver 를 넘겨줘서 로그인 한 뒤 다시 driver 넘겨받기
			driver := return_NewWindowDriver_After_FGLogin(driver)
			
			; 이렇게 일단 두 번째 로그인한 창을 닫아야 비로서 첫 번째 프로파일 저장한 경로가 완성(?) 되어 첫 번째 프로파일 저장한 프로파일 경로를 이용할 수 있다.
			driver.close()
			
			
				
			; ########## 세번째 창 열기 (실제로 열고 싶은 URL 열기) ##########

			driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
			driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
			driver.AddArgument("--start-maximized") ; 창 최대화 하기
			

			; chrome_Profile.txt 파일 내용을 읽어서 chrome_Profile 변수에 저장하기
			FileRead, chrome_Profile, %A_ScriptDir%\CreatedFiles\NewWindowProcessing\chrome_Profile.txt

			; 프로파일 경로 읽기. 두 번째 로그인 해서 프로파일 경로를 완성(?) 시켜 줬기 때문에 지금부터는 프로파일 경로를 제대로 읽는다		
			driver.SetProfile(chrome_Profile) ; chrome_Profile 변수를 읽어서 처음 로그인 한 프로파일 경로를 읽는다


			; 이제는 로그인 화면으로 넘어가지 않고 제대로 뉴오더 화면으로 곧장 이동한다
			driver.Get(url)
			
;MsgBox, 내가 가고 싶은 페이지. 로그인 한 뒤 열렸음
				
			
		} ; if 닫기 - FG 로그인 페이지라면
		
	} ; if 닫기 - 만약 현재 페이지가 FG 페이지라면
	
	
	; 만약 현재 페이지가 LAS 페이지라면
	else if(RegExMatch(driver.Url, "imU)lashowroom")){
		
		; LAS 로그인 페이지라면
		if(RegExMatch(driver.Url, "imU)login")){


			; LAS 로그인만 하는 메소로 driver 넘겨서 로그인 후 driver 다시 받아오기
			driver := LAS_LoginOnly(driver)


			; 원하는 URL로 이동
			driver.Get(url)
			
			
			
			
			
			
		} ; if 닫기 - LAS 로그인 페이지라면	
	} ; if 닫기 - 만약 현재 페이지가 LAS 페이지라면
	
	

	
	
	
;MsgBox, 내가 가고 싶은 페이지. 로그인 하기 전에 열렸음


	; driver 리턴하기
	return driver
	
	
	
	
} ; goToURl_AfterLogIn_IfNeeded(URL) 메소드 끝
















	; ############################################################################################
	; FG 열고 싶은 창을 로그인 없이 열 수 있게끔 로그인 두 번 해서 프로필 파일 저장한 뒤 driver 리턴하기
	; ############################################################################################
	return_NewWindowDriver_After_FGLogin(driver){
		
		; ########## 첫번째 로그인 하기 ##########
		; driver 를 넘겨줘서 로그인 한 뒤 다시 driver 넘겨받기
		driver := FG_LoginOnly(driver)
		
		
		; 프로파일 경로 얻는 화면으로 이동
		driver.Get("chrome://version/")

		; 프로파일 경로 찾아서 chrome_Profile 변수에 넣기
		Xpath = //*[@*='profile_path']
		chrome_Profile := driver.FindElementByXPath(Xpath).Attribute("innerText")

		; chrome_Profile.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
		FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\NewWindowProcessing\chrome_Profile.txt, 1
			
		; chrome_Profile 변수 안에 있는 메모 내용 chrome_Profile.txt 파일에 저장하기
		FileAppend, %chrome_Profile%, %A_ScriptDir%\CreatedFiles\NewWindowProcessing\chrome_Profile.txt

		; 프로파일 경로 확인
;MsgBox, % "chrome_Profile : " . chrome_Profile

		; 일단 창을 닫는다
		driver.close()
		




		; ########## 두번째 로그인 하기 ##########
		; 창을 연 뒤 첫번째 로그인 한 뒤 얻은 프로파일 경로를 읽는다
		; 그래도 로그인을 하라고 하기 때문에 또 로그인을 하게 되면 프로파일 경로는 첫 번째 프로파일 경로와 똑같고 맨 마지막에 \Default 이 붙는다.
		; 하지만 실제로 사용할 것은 첫 번째 프로파일 경로이다.

		driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
		driver.addargument("--start-minimized") ; 창 최소화 하기
		;~ driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
		;~ driver.AddArgument("--start-maximized") ; 창 최대화 하기
		 
		 
		; 프로파일 경로 읽기
		driver.SetProfile(chrome_Profile) ; 이렇게 chrome_Profile 변수를 읽어서 처음 로그인 한 프로파일을 읽는다


		; url로 이동해 보기. 하지만 로그인 화면으로 자동으로 넘어가게 된다
		driver.Get(url)
		
		
		; 현재 페이지가 로그인 페이지라면 실행
		if(RegExMatch(driver.Url, "imU)login")){
			

			
			; 두 번째 로그인 하기
			; driver 를 넘겨줘서 로그인 한 뒤 다시 driver 넘겨받기
			driver := FG_LoginOnly(driver)

				
			; 프로파일 경로 얻는 화면으로 이동
			driver.Get("chrome://version/")

			; 프로파일 경로 찾아서 2NDchrome_Profile 변수에 넣기
			Xpath = //*[@*='profile_path']
			2NDchrome_Profile := driver.FindElementByXPath(Xpath).Attribute("innerText")

			; 프로파일 경로 확인
;MsgBox, % "2NDchrome_Profile : " . 2NDchrome_Profile
		}


		; 이렇게 두 번째 로그인 한 창을 닫아야 첫 번째 프로파일 저장한 프로파일 경로를 이용할 수 있다.
		;~ driver.close()
		
		; 일단 driver 넘겨준다.호출한 곳에서 창을 닫아야 프로파일 경로가 완성(?) 된다
		return driver
	}




	
	; FG 로그인만 하는 메소드
	FG_LoginOnly(driver){


		; ############
		; 천희 ID & PW
		; ############
		Loginname = customer3
		Password = Jo123456789


		; 아이디 입력
		Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[1]/input
		Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[1]/input[2]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Loginname) ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A 한 뒤 로그인 입력


		; 비밀번호 입력
		Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[2]/input
		Xpath = /html/body/fg-root/div[1]/fg-public-layout/fg-auth/div[1]/div/div/div[1]/div/div/form/div[2]/input[2]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Password).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 비밀번호 입력 후 엔터쳐서 로그인하기

		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		
		return driver
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	; LAS 로그만 하는 메소드
	LAS_LoginOnly(driver){
		
		; ######
		; 천희
		; ######			
		Loginname = jodifl3
		Password = jodifl
		SecurityCode = 7864
		CCPWD = FC83D28D
		

		; 아이디 입력
		Xpath = //*[@id="uname"]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Loginname) ; 기존에 혹시 자동완성으로 정보가 채워져 있다면 지우기 위해 Ctrl+A 한 뒤 로그인 입력


		; 비밀번호 입력
		Xpath = //*[@id="login_pwd"]
		driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(Password).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 비밀번호 입력 후 엔터쳐서 로그인하기


		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		

		; 현재 url이 Security Verification Code 입력 화면이면 보안코드 입력하기
		; https://admin.lashowroom.com/login_verify.php
		if(RegExMatch(driver.Url, "imU)login_verify")){
			
			; Security Code 입력
			Xpath = //*[@id="verification_code"]
			driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(SecurityCode).sendKeys(driver.Keys.ENTER) ; Ctrl+A 한 뒤 SecurityCode 입력 후 엔터쳐서 로그인하기				
			Sleep 500

			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
			
			return driver			
		}
	}
	
	
	
	
	

	
	
; ############################################################################################################################################################################################################################################################################
; ###########################################################################################################     넘겨받은 URL 로 이동하기 끝   ################################################################################################################################
; ############################################################################################################################################################################################################################################################################































; ############################################################################################################################################################################################################################################################################
; ##################################################################################################################     LAS 처리   ##########################################################################################################################################
	
	
	
	
	; ###########################################################
	; 뉴오더가 아닌(예를 들어 allocation) LAS 처리
	; 해당 PONumber 검색해서 정보 읽어온 뒤 가져온 정보와 driver 리턴
	; ###########################################################
	processLAS_which_from_Not_New_Orders(driver, PONumber){
		
		
		Arr_Memo := object()
		Arr_CC := object()
		Arr_BillingADD := object()
		Arr_ShippingADD := object()
		
		Arr_ShippingOptionStatus := object()
				
		
		
		URL = https://admin.lashowroom.com/orders_cur_month.php ; 오더 검색창 주소
		driver := goToURl_AfterLogIn_IfNeeded(driver, URL) ; url 로 이동. 필요하다면 로그인 (LAS 는 창 열때마다 무조건 로그인 함)
		
		
		; 검색란에 PO Number 입력하기
		Xpath = //*[@id="search_po"]
		;~ driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(PONumber).sendKeys(driver.Keys.ENTER)
		driver.FindElementByXPath(Xpath).SendKeys(PONumber).sendKeys(driver.Keys.ENTER)


	
		; PO Number 링크 나타날때까지 기다림
		;~ Xpath = //*[text() = '%PONumber%']
		Xpath = //*[contains(text(), '%PONumber%')]
		Sleep 500
		while(!driver.FindElementByXPath(Xpath).Attribute("innerText"))
			Sleep 100
		
		;~ Sleep 1000
		


		; 검색된 PO 결과 창에서 가장 처음 값을 클릭하기 (토씨 하나 안 틀린 PO 값을 클릭하는게 아니라 PO 번호가 포함된 가장 처음의 링크 클릭)
		; 이게 희안한게 브라우저 글자 크기를 확대하면 작동을 안 한다
		if(driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]")){

		
			; 새 창이 아닌 새 탭에서 열기 위해 Ctrl 누르고 그리로 이동하기 위해 Shift 누른다. (키보드로 하려면 두 키 동시에 누르고 클릭하면 됨)
			driver.sendKeys(driver.Keys.CONTROL)
			driver.sendKeys(driver.Keys.SHIFT)

			; 링크 클릭
			driver.FindElementByXPath("(//*[contains(text(), '" PONumber "')])[1]").click()
			
			
			; 새 탭을 열기 위해 눌렀던 Ctrl 과 Shift 키 누른것 해제하기 위해
			driver.sendKeys(driver.Keys.CONTROL)
			driver.sendKeys(driver.Keys.SHIFT)			
			
			
			
			driver.SwitchToNextWindow() ; 새로 연 탭으로 콘트롤 옮김
			
/*			
			driver.waitForPageToLoad() ; 페이지 로딩하는 동안 기다림
			
			;~ MsgBox, 콘트롤 옮겼음
			
			driver.Get("http://www.google.com")
			
			MsgBox, 새로 열린 탭이 구글로 바뀌었어야 됨
*/			
			


			; Shipping Method
			; Shipping Option 의 드롭박스의 상태값을 읽어서 변수에 저장
			Shipping_Method_Xpath = //*[@id="id_shipmode"]
			ShippingMethodStatus := driver.FindElementByXPath(Shipping_Method_Xpath).Attribute("value")
			
			
;			MsgBox, % ShippingMethodStatus
			
			
			; UPSG 가 아니라면
			if(ShippingMethodStatus != "UPS Ground"){
				
				; LAS consolidation 인지 확인
				if(ShippingMethodStatus == "LAS Order Consolidation"){
;					MsgBox, It's LAS Consolidation

					; ShippingOptionStatus 변수 값을 3으로 바꿈(LAS Consolidation 위함)
					Arr_ShippingOptionStatus[1] := "3"
					
				}
				
				; UPSG 가 아니면서 LAS Conslidation 도 아닌 주문들 처리
				else
				{	
;					MsgBox, It's neither UPSG nor LAS consolidation
					
					; ShippingOptionStatus 변수 값을 2로(UPSG 도 아니고 LAS consolidation 도 아닌 상태. 'delivery, 2nd, 3rd day USPS' 등을 위해) 바꿈
					Arr_ShippingOptionStatus[1] := "2"
					
				} ; end of else - ; UPSG 가 아니면서 LAS Conslidation 도 아닌 주문들 처리

			}
			
			; UPSG 주문일 때
			else{
				
;				MsgBox, It's UPSG
			
				; ShippingOptionStatus 변수 값을 1로 바꿈(UPSG 위함)
				Arr_ShippingOptionStatus[1] := "1"
			}
			
			
;			MsgBox, % Arr_ShippingOptionStatus[1]








			
			; Buyer Notes
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[3]/td[1]
			Arr_Memo[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
			Data := Arr_Memo[1]
			Data := RegExReplace(Data, "Comment: None(.*)", "$1")  ; $1 역참조를 사용하여 Comment: None 이외의 메모 내용이 있으면 변수에 저장
			StringUpper, Data, Data ; 대문자로 바꾸기
			Arr_Memo[1] := Data
			
			
			; Contact Name on Billing Add
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[9]/td[2]
			Arr_CC[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")				
			
			

			; Billing Add
			; ADD1
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[2]
			Arr_BillingADD[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")
			Data := Arr_BillingADD[1]
			StringUpper, Data, Data ; 대문자로 바꾸기
			Arr_BillingADD[1] := Data
			
			; ADD2
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]
			Arr_BillingADD[2] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
			Data := Arr_BillingADD[2]
			StringUpper, Data, Data ; 대문자로 바꾸기
			Arr_BillingADD[2] := Data
			
			; CITY
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[2]
			Arr_BillingADD[3] := driver.FindElementByXPath(Xpath).Attribute("textContent")
			Data := Arr_BillingADD[3]
			StringUpper, Data, Data ; 대문자로 바꾸기
			Arr_BillingADD[3] := Data
					
			; STATE
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[5]/td[2]
			Arr_BillingADD[4] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
			Data := Arr_BillingADD[4]
			StringUpper, Data, Data ; 대문자로 바꾸기
			Arr_BillingADD[4] := Data
			
			; ZIP
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[6]/td[2]
			Arr_BillingADD[5] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
			Data := Arr_BillingADD[5]
			StringUpper, Data, Data ; 대문자로 바꾸기
			Arr_BillingADD[5] := Data
					
			; COUNTRY
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[7]/td[2]
			Arr_BillingADD[6] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
			Data := Arr_BillingADD[6]
			StringUpper, Data, Data ; 대문자로 바꾸기
			Arr_BillingADD[6] := Data
					
			; PHONE
			Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td[2]
			Arr_BillingADD[7] := driver.FindElementByXPath(Xpath).Attribute("textContent")
			Data := Arr_BillingADD[7]
			Data := RegExReplace(Data, "[^0-9]", "") ;숫자만 저장
			Arr_BillingADD[7] := Data
			
			
			


			; LAS consolidation 이 아닐때만 배송 주소 정보 저장하기
			; consolidation 일때는 배송 주소를 읽을 필요가 없으니까. 그리고 Xpath 가 바뀌어서 읽으려고 해도 에러남
			if(Arr_ShippingOptionStatus[1] != "3"){
				
				; Shipping Add
				; ADD1
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[2]/td[2]
				Arr_ShippingADD[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")
				Data := Arr_ShippingADD[1]
				StringUpper, Data, Data ; ???? ???
				Arr_ShippingADD[1] := Data
				
				; ADD2
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[3]/td[2]
				Arr_ShippingADD[2] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
				Data := Arr_ShippingADD[2]
				StringUpper, Data, Data ; ???? ???
				Arr_ShippingADD[2] := Data
				
				; CITY
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[4]/td[2]
				Arr_ShippingADD[3] := driver.FindElementByXPath(Xpath).Attribute("textContent")
				Data := Arr_ShippingADD[3]
				StringUpper, Data, Data ; ???? ???
				Arr_ShippingADD[3] := Data
						
				; STATE
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[5]/td[2]
				Arr_ShippingADD[4] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
				Data := Arr_ShippingADD[4]
				StringUpper, Data, Data ; ???? ???
				Arr_ShippingADD[4] := Data
				
				; ZIP
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]
				Arr_ShippingADD[5] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
				Data := Arr_ShippingADD[5]
				StringUpper, Data, Data ; ???? ???
				Arr_ShippingADD[5] := Data
						
				; COUNTRY
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[7]/td[2]
				Arr_ShippingADD[6] := driver.FindElementByXPath(Xpath).Attribute("textContent")		
				Data := Arr_ShippingADD[6]
				StringUpper, Data, Data ; ???? ???
				Arr_ShippingADD[6] := Data
						
				; PHONE
				Xpath = //*[@id="orderedit_form"]/div/div[7]/table/tbody/tr[1]/td[2]/table/tbody/tr[10]/td[2]
				Arr_ShippingADD[7] := driver.FindElementByXPath(Xpath).Attribute("textContent")
				Data := Arr_ShippingADD[7]
				Data := RegExReplace(Data, "[^0-9]", "") ;??? ??
				Arr_ShippingADD[7] := Data

			} ; if ends - if(Arr_ShippingOptionStatus[1] != "3")

			



	;		MsgBox, % Arr_Memo[1] . "`n`n" . Arr_CC[1]
					
			/* 배열로부터 읽기 첫 번째 방법
			Loop % Arr_ShippingADD.Maxindex(){
				MsgBox % "Element number " . A_Index . " is " . Arr_ShippingADD[A_Index]
			}
			*/




			


			
			; Update 버튼 클릭하기			
			Update_Xpath = //*[@id="update_order"]
			driver.FindElementByXPath(Update_Xpath).click()


			
			
			
			return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo, Arr_ShippingOptionStatus]
		
		
		
					
			

		}
		
		
		
		return driver
		
		
	} ; GetInfoFromLASPage(PONumber) 메소드 끝	
	
	
	
	
; ##################################################################################################################     LAS 처리 끝  ##########################################################################################################################################		
; #############################################################################################################################################################################################################################################################################
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
; #############################################################################################################################################################################################################################################################################
; ##################################################################################################################   FG 처리 시작 ###########################################################################################################################################


; 주문 페이지의 정보 읽어서 리턴해주기
getInfoOnFG_And_Return_That(driver, CustomerPO, IsItFromNewOrder, IsItFromExcelFile){
;~ GettingInfoFromCurrentPage(CustomerPO, IsItFromNewOrder, IsItFromExcelFile){

	Arr_CC := object()
	Arr_Memo := object()
	
	
	; Shipping Method 상태 알아내기
	; UPS Ground 이면 값은 3
	; 값이 3이 아니면 UPS Ground 가 아님
	Shipping_Method_Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[2]/div[2]/ul/li[1]/span/span/select
	ShippingMethodStatus := driver.FindElementByXPath(Shipping_Method_Xpath).Attribute("value")
	
;MsgBox, % "ShippingMethodStatus : " . ShippingMethodStatus
	
	
	; Name
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[1]/ul/li[2]/span[2]
	Name := driver.FindElementByXPath(Xpath).Attribute("innerText")
	StringUpper, Name, Name ; 대문자로 바꾸기
	Arr_CC[1] := Name		
		
	; Buyer Notes	
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[1]
	Arr_Memo[1] := driver.FindElementByXPath(Xpath).Attribute("textContent")
	Data := Arr_Memo[1]
	StringUpper, Data, Data
	Arr_Memo[1] := Data
		
	; Additional Info	
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[1]/div/textarea[2]
	Arr_Memo[2] := driver.FindElementByXPath(Xpath).Attribute("textContent")
	Data := Arr_Memo[2]
	StringUpper, Data, Data
	Arr_Memo[2] := Data
		
	; Staff Notes (Internal Use Only)
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/div/div/div[2]/div/textarea
	Arr_Memo[3] := driver.FindElementByXPath(Xpath).Attribute("value")
	Data := Arr_Memo[3]
	StringUpper, Data, Data
	Arr_Memo[3] := Data
	
	
	
	

		
	; Billing Address		
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[1]/ul/li[3]/span[2]
	BillingAdd := driver.FindElementByXPath(Xpath).Attribute("innerText")

	; Shipping Address		
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[2]/ul/li[3]/span[2]
	ShippingAdd := driver.FindElementByXPath(Xpath).Attribute("innerText")		

	Arr_BillingADD := StrSplit(BillingAdd, ",")  ; 콤마 나올때마다 문자열 나눠서 Arr_BillingADD 배열에 넣기
	Arr_ShippingADD := StrSplit(ShippingAdd, ",")  ; 콤마 나올때마다 문자열 나눠서 Arr_BillingADD 배열에 넣기
		

	; Arr_BillingADD2 찾아서 배열 5번째에 넣고 1번째에는 ADD1만 남기기
	UnquotedOutputVar = im)(( unit| Suite| Ste| #| Apt| SPACE| BLDG| Building| Sujite| Sujite).*)
	;~ UnquotedOutputVar = im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite|Sujite).*)
	;~ UnquotedOutputVar = im)((\sunit|\sSuite|\sSte|\s#|\sApt|\sSPACE|\sBLDG|\sBuilding|\sSujite|\sSujite).*)
	Arr_BillingADD[5] := M_driver.FindAdd2_In_Add1(Arr_BillingADD[1], UnquotedOutputVar) ; Arr_BillingADD[1] 에 들어있는 전체 주소를 넘겨서 ADD2 만 Arr_BillingADD[5] 에 넣기
	Arr_BillingADD[1] := M_driver.DeleteAdd2_In_Add1(Arr_BillingADD[1], UnquotedOutputVar) ; 전체 주소 중 ADD2를 지운뒤 Arr_BillingADD[1]에 넣기
		
;	MsgBox, % "add1 : " Arr_BillingADD[1] . "`n" . "add2 : " . Arr_BillingADD[5]


	; ZIP 찾아서 배열 6번째에 넣고 3번째에는 State(州)만 넣기
	UnquotedOutputVar = im)(\d.*)
	Arr_BillingADD[6] := M_driver.FindAdd2_In_Add1(Arr_BillingADD[3], UnquotedOutputVar) ; Arr_BillingADD[3] 에 있는 State(州) + Zip 을 Zip만 Arr_BillingADD[6] 에 넣기
	Arr_BillingADD[3] := M_driver.DeleteAdd2_In_Add1(Arr_BillingADD[3], UnquotedOutputVar) ; Arr_BillingADD[3] 값에서 Zip 지운 뒤 State(州) 만 Arr_BillingADD[3] 에 넣기
		
		
;		MsgBox, % "State(州) : " Arr_BillingADD[3] . "`n" . "Zip : " . Arr_BillingADD[6]







	; Arr_ShippingADD2 찾아서 배열 5번째에 넣고 1번째에는 ADD1만 남기기
	UnquotedOutputVar = im)(( unit| Suite| Ste| #| Apt| SPACE| BLDG| Building| Sujite| Sujite).*)
	;~ UnquotedOutputVar = im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite|Sujite).*)
	;~ UnquotedOutputVar = im)((\sunit|\sSuite|\sSte|\s#|\sApt|\sSPACE|\sBLDG|\sBuilding|\sSujite|\sSujite).*)
	Arr_ShippingADD[5] := M_driver.FindAdd2_In_Add1(Arr_ShippingADD[1], UnquotedOutputVar) ; Arr_ShippingADD[1] 에 들어있는 전체 주소를 넘겨서 ADD2 만 Arr_ShippingADD[5] 에 넣기
	Arr_ShippingADD[1] := M_driver.DeleteAdd2_In_Add1(Arr_ShippingADD[1], UnquotedOutputVar) ; 전체 주소 중 ADD2를 지운뒤 Arr_BillingADD[1]에 넣기


	; ZIP 찾아서 배열 6번째에 넣고 3번째에는 State(州)만 넣기
	UnquotedOutputVar = im)(\d.*)
	Arr_ShippingADD[6] := M_driver.FindAdd2_In_Add1(Arr_ShippingADD[3], UnquotedOutputVar) ; Arr_ShippingADD[3] 에 있는 State(州) + Zip 을 Zip만 Arr_ShippingADD[6] 에 넣기
	Arr_ShippingADD[3] := M_driver.DeleteAdd2_In_Add1(Arr_ShippingADD[3], UnquotedOutputVar) ; Arr_ShippingADD[3] 값에서 Zip 지운 뒤 State(州) 만 Arr_ShippingADD[3] 에 넣기
		
		
	; Phone Number
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[3]/div[2]/div[2]/div[1]/ul/li[4]/span[1]
	PhoneNumber := driver.FindElementByXPath(Xpath).Attribute("innerText")
		
	PhoneNumber := RegExReplace(PhoneNumber, "[^0-9]", "") ;숫자만 저장
	Arr_BillingADD[7] := PhoneNumber


	; 카드 정보 순서 변경하기
	Arr_Temp := Arr_BillingADD.Clone()		
		
	Arr_BillingADD[1] := Arr_Temp[1] ; ADD1
	Arr_BillingADD[2] := Arr_Temp[5] ; ADD2
	Arr_BillingADD[3] := Arr_Temp[2] ; CITY
	Arr_BillingADD[4] := Arr_Temp[3] ; STATE
	Arr_BillingADD[5] := Arr_Temp[6] ; ZIP
	Arr_BillingADD[6] := Arr_Temp[4] ; COUNTRY
	Arr_BillingADD[7] := Arr_Temp[7] ; PHONE
		

	Arr_Temp := []
	Arr_Temp := Arr_ShippingADD.Clone()
		
		
	Arr_ShippingADD[1] := Arr_Temp[1]
	Arr_ShippingADD[2] := Arr_Temp[5]
	Arr_ShippingADD[3] := Arr_Temp[2]
	Arr_ShippingADD[4] := Arr_Temp[3]
	Arr_ShippingADD[5] := Arr_Temp[6]
	Arr_ShippingADD[6] := Arr_Temp[4]
	Arr_ShippingADD[7] := Arr_Temp[7]
		



		/*
			Arr_BillingADD[1] -> ADD1
			Arr_BillingADD[2] -> City
			Arr_BillingADD[3] -> State
			Arr_BillingADD[4] -> United States
			Arr_BillingADD[5] -> ADD2
			Arr_BillingADD[6] -> Zip
			Arr_BillingADD[7] -> PhoneNumber
		*/
		
		
	; Arr_BillingADD 배열의 모든 값을 대문자로 바꾸고 양쪽에 공란 없애기
	Loop % Arr_BillingADD.Maxindex(){

		Data := Arr_BillingADD[A_Index]
		
		; ADD2 는 대문자로 바꾸지 않기 위해. #를 StringUpper로 처리하면 에러 난다. ADD2 에서 #를 지운다
		if(A_Index == 2){
			StringReplace, Data, Data, #, , All
			Arr_BillingADD[A_Index] := Data
			continue
		}
						
		StringUpper, Data, Data
		Arr_BillingADD[A_Index] := Trim(Data)
		;~ MsgBox % "Element number " . A_Index . " is " . Arr_BillingADD[A_Index]
	}


	; Arr_ShippingADD 배열의 모든 값을 대문자로 바꾸고 양쪽에 공란 없애기
	Loop % Arr_ShippingADD.Maxindex(){
			
		Data := Arr_ShippingADD[A_Index]
		
		; ADD2 는 대문자로 바꾸지 않기 위해. #를 StringUpper로 처리하면 에러 난다. ADD2 에서 #를 지운다
		if(A_Index == 2){
			StringReplace, Data, Data, #, , All
			Arr_ShippingADD[A_Index] := Data
			continue
		}
			
		StringUpper, Data, Data
		Arr_ShippingADD[A_Index] := Trim(Data)
		;~ MsgBox % "Element number " . A_Index . " is " . Arr_ShippingADD[A_Index]
	}



		


		return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo, ShippingMethodStatus]



	

	MsgBox, 어디까지 진행됐나	
	
	return driver
	
} ; getInfoOnFG_And_Return_That 메소드 끝






; ########################
; 현재 페이지의 Order Status 가 New Orders 이거나 Back Ordered 일때 Confirmed Orders 로 바꾸기
; ########################
changeNewOrders_To_ConfirmedOrders(driver){
	
	
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div/ul/li[2]/span/span[1]/select
	;~ HasDropdownMenuChanged := FG.compareTheValueOfDropdownMenuAndChangeTheStatusToPreferenceOne(Xpath, 2, "Confirmed Orders")
	CurrentStatus := driver.FindElementByXPath(Xpath).Attribute("value") ; 상태 값을 얻기
	
;MsgBox, % CurrentStatus	
	
	; New Orders 일때만 Confirmed Orders 로 바꾸기
	if(CurrentStatus == 1||CurrentStatus == 7){
		
		; Confirmed Orders 로 드롭박스 바꾸기
		driver.FindElementByXPath(Xpath).SendKeys("Confirmed Orders")
		
		; Update 버튼 클릭
		Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[2]/div[2]/div[1]/div/ul/li[2]/span/button[1]
		driver.FindElementByXPath(Xpath).click()
	}
	
	
	return driver

} ; changeNewOrders_To_ConfirmedOrders 메소드 끝









; ########################
; 새 탭 열기
; 가장 위에 있는 PO 번호 클릭해서 열기
; ########################
openNewTab_clickMostTopPO#(driver, CustomerPO){
	

	Sleep 2000
	
	; 화면에서 Customer PO 에 있는 outer HTML 값이 나올때까지 루프 반복
	; 즉, 화면 로딩이 끝날때까지 루프 반복
	; 그런데 이거 안 먹힘. 젠장. 로딩하는 동안 기다리는거 어떻게 구현해야 되는거야?
	
	outerHTMLofTheCustomerPO = 0
	Loop{
		
		; 화면에서 Customer PO 에 있는 outer HTML 값 outerHTMLofTheCustomerPO 에 저장하기
		outerHTMLofTheCustomerPO := driver.FindElementByXPath("(//*[contains(text(), '" CustomerPO "')])[1]").Attribute("outerHTML")
		Sleep 100
		
		if(outerHTMLofTheCustomerPO){
			break
		}
	}


	; outerHTML 에서 url 페이지 들어가기 위해 Customer PO 의 고유 아이디만 uniqueIDofCustPO 변수에 저장
	UnquotedOutputVar = imU)href="#/order/(.*)">
	RegExMatch(outerHTMLofTheCustomerPO, UnquotedOutputVar, SubPat)



;	MsgBox, % "THE UNIQUE NUMBER OF CUST'S ORDER PAGE ID IS : " . SubPat1



	; 위의 동작에서 얻은 고유 아이디에 기본 url 주소를 붙이면 해당 Customer PO 의 주문 페이지가 된다.
	; 그곳으로 이동한다
	; SubPat1 를 찾았을 때만 이동한다
	if(SubPat1){
		
		URLofCustPO := "https://vendoradmin.fashiongo.net/#/order/" . SubPat1
		driver.executeScript("window.open('" URLofCustPO "')") ; 새 탭 열기 - 지정된 url 열기
	}
	; 가끔 outerHTMLofTheCustomerPO 읽을때 
	; <a _ngcontent-c10="" href="#/order/13408498">MTR1F38167D12</a>
	; 이런 형식이 아닌 다른 이상한? 값의 형태가 있을때가 있다.
	; 그런 값에서는 RegExMatch 이용해서 주문페이지의 고유 아이디를 추출할 수 없다
	; 또한 PONumber 에 해당하는 링크를 클릭할 수도 없다
	; 왜 그런지 모르겠다. 그냥 에러 메세지 띄우고 직접 클릭하라고 하는게 지금까지 찾아낸 유일한 해법이다.
					
	; MTR1F38167D12 <- 이게 클릭이 안 된다

	else if(!SubPat1){
		SoundPlay, %A_WinDir%\Media\Ring02.wav
		MsgBox, 262144, Title, AN ERROR OCCURRED. PLEASE CLICK THE CUSTOMER CODE BELOW MANUALLY.`n`n`n%PONumber%
		;~ driver.FindElementByXPath("//*[contains(text(), '" PONumber "')]").click()
		;~ driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
		;~ driver := ChromeGet()
	}	

	
	
	driver.SwitchToNextWindow ; 새로 연 탭으로 콘트롤 옮김



;~ /*					
	; 오더 페이지에 제대로 들어갔는지 확인하기
	; 현재 열린 창과 고객의 주문 페이지가 맞지 않으면
	; 현재 창 리프레쉬 한 뒤 continue 로 루프 다시 시작해보기
	; SubPat1 를 찾았을 때만 제대로 들어갔는지 확인한다
	; 찾지 못했을 땐 위에서 그냥 수동으로 클릭했다.
	if(SubPat1){
		Sleep 1000
		CurrentURL := driver.Url
		
		if(CurrentURL != URLofCustPO)
		{			
			MsgBox, 262144, IT'S NOT CUSTOMER'S ORDER PAGE`nRESTART LOADING THE ORDER PAGE AGAIN`n`nCurrentURL : `n%CurrentURL%`nURLofCustPO : `n%URLofCustPO%
			
			; 아래 코드 사용해서 크롬창 닫고 아예 처음부터 다시 시작하기
			driver.close() ; closing just one tab of the browser
			
			; 전체 오더 검색창 주소로 이동하기
			URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
			driver := goToURl_AfterLogIn_IfNeeded(driver, URL) ; 원하는 url로 이동

			; 전체 오더 검색창 주소로 이동한 뒤
			; 검색조건을 PO 번호로 바꾼 뒤 PO 번호로 찾기
			driver := findOrdersByPO#(driver, CustomerPO)
			
			; 가장 위에 있는 PO 번호를 새탭으로 열기
			driver := openNewTab_clickMostTopPO#(driver, CustomerPO)
			
			; 이 메소드 재귀호출하기
			getInfoOnFG_And_Return_That(driver, CustomerPO, IsItFromNewOrder, IsItFromExcelFile)
		}
	}

*/



	; 오더 페이지 제대로 들어갔으니 리턴으로 메소드 끝내기
	return driver

	
} ; openNewTab_clickMostTopPO#() 메소드 끝






; ########################
; 검색 옵션을 PO Number 로 바꾼 뒤 PO 번호 입력해서 검색하기
; ########################	
findOrdersByPO#(driver, CustomerPO){
	
	; CUSTOMER PO 검색칸에 PO Number 입력
	Xpath = (/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[3]/input)[2]
	driver.FindElementByXPath(Xpath).sendKeys(driver.Keys.CONTROL, "a").SendKeys(CustomerPO)
	

	; PO Number 로 바꾸기
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[2]/select
	driver := changeDropboxStatus(driver, Xpath, po, "PO Number") ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기


	; Input Period 상태를 Last 365 Days 로 바꾸기
	Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-orders/fg-order-search/div/div[2]/div/div/div[1]/select
	driver := changeDropboxStatus(driver, Xpath, 33, "Last 365 Days") ; Xpath, StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)을 parameter로 넘겨주기

	
	return driver
	
} ; findOrdersByPO#(driver, CustomerPO) 메소드 끝








; ########################
; DropDown 메뉴 값을 읽은 후 argument 로 받은 값이 아니면 받은 값으로 바꾼뒤 driver 넘겨주기
; StatusToBePreferred(바뀌었으면 하는 상태 값), PreferredStatusValue(바꿀 상태값)
; ########################	
changeDropboxStatus(driver, Xpath, StatusToBePreferred, PreferredStatusValue){
		
	
	; Xpath 에 있는 element 가 활성화 되어 수정 가능한지 확인하기. 0이 반환되면 사용 불가
	IsItEnabled := driver.FindElementByXPath(Xpath).isEnabled()
		
	if(IsItEnabled == 0){
		MsgBox, It's Disabled.
	}
		
	; DropDown Manu 상태를 읽은 후 원하는 상태가 아니면 원하는 상태로 바꾸기
	CurrentStatus := driver.FindElementByXPath(Xpath).Attribute("value") ; 상태 값을 얻기
		
	;~ MsgBox, % CurrentStatus
		
	if(CurrentStatus != StatusToBePreferred){ ; 상태값이 원하는 값이 아니면 if문으로 들어가서 원하는 상태로 바꾸기
		driver.FindElementByXPath(Xpath).SendKeys(PreferredStatusValue)
	}


	; driver 리턴해주기
	return driver
} ; changeDropboxStatus(driver, Xpath, StatusToBePreferred, PreferredStatusValue) 메소드 끝







; ##################################################################################################################   FG 처리 끝 #############################################################################################################################################
; #############################################################################################################################################################################################################################################################################
	
		
	
	