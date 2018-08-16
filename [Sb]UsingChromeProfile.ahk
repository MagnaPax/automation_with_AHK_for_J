;###################################################################
; 크롬에 저장된 쿠키 이용하기(로그인 한 뒤 다시 로그인 할 필요 없게 하기)
;###################################################################



; put this "chrome://version/"  in the url in chrome to find path to profile
 ;https://stackoverflow.com/questions/25779027/load-default-chrome-profile-with-webdriverjs-selenium
 
 /* 
 driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
; put this "chrome://version/"  in the url in chrome to find path to profile
 ;https://stackoverflow.com/questions/25779027/load-default-chrome-profile-with-webdriverjs-selenium
driver.SetProfile("H:\Temp\Chrome\Cache\cache\Default") ; 'Full path of the profile directory
url:="https://www.linkedin.com/feed/?trk="
driver.Get(url) 
 */
 
 
 
 

/*
이게 조금 재미있는것이, 로그인을 두 번 한 뒤 첫 로그인 한 뒤의 profile 경로를 대입해야지만 제대로 작동한다. 무슨 뜻이냐면

1. 창을 열고 로그인을 한다
	1-1. 프로파일 창을 열어서 경로를 확인한다면 C:~~~~\Default 이렇게 끝이 나는 것을 확인할 수 있다.
	1-2. 프로파일 경로를 chrome_Profile 변수에 저장한다
	1-3. 이렇게 로그인 된 창을 닫는다

2. 창을 새로 연다
	2-1. 이때는 창을 연 뒤 driver.SetProfile(chrome_Profile) 이용하여 chrome_Profile 에 저장된 프로파일을 읽는다
	2-2. 이렇게 읽었어도 여전히 프로파일을 못 읽고 다시 로그인 하라고 뜰 것이다
	2-3. 로그인 한 뒤 프로파일 창을 열어서 경로를 확인한다면 위의 경로와 똑같은데 대신 맨 마지막에 Default 하나가 더 붙어서 다음과 같이 된다 C:~~~~\Default\Default
	2-4. 로그인 한 뒤 창을 닫는다.

3. 창을 새로 열어서 처음에 로그인 한 뒤 저장한 프로파일 경로 chrome_Profile를 읽으면 이때는 다시 로그인을 묻지 않고 제대로 프로파일 읽고 작동한다
	3-1. 실전에서는 chrome_Profile 를 파일로 저장한 뒤 읽으면 될 것 같다



** 첫 번째 로그인 만으로는 프로파일이 완성되지 못하는 것 같다. 꼭 두 번째로 새 창을 연 뒤 첫 번째 프로파일을 읽은뒤 다시 로그인 해서 첫 번째 프로파일 경로 끝에 \Default를 붙여야만 비로소 첫 번째 프로파일 경로가 완성되는 것 같다.
** 여기서 주의할 점은 비록 두 번째 로그인 해서 C:~~~~\Default\Default 이렇게 끝나는 프로파일 경로를 얻었지만 이후로 실제 사용할 프로파일 경로는 첫 번째 로그인 후 나온 C:~~~~\Default 라는 사실이다.



** 프로파일을 읽는 중요 코드는 아래와 같은데 둘 다 작동한다. 세 번째는 변수를 이용해 프로파일 읽는 코드이다.

; 직접 프로파일 경로를 입력하여 얻기 방법 1
driver.SetProfile("C:\Users\JODIFL4\AppData\Local\Temp\Selenium\scoped_dir2036_26699\Default")

; 직접 프로파일 경로를 입력하여 얻기 방법 2
driver.AddArgument("--user-data-dir=C:\Users\Hahn\AppData\Local\Temp\Selenium\scoped_dir7552_19270\Default")

; chrome_Profile 변수를 통해 프로파일을 읽는다. 
driver.SetProfile(chrome_Profile)



*/












	; ########## 첫번째 창 열기 ##########
	; 로그인 뒤 프로파일 경로를 chrome_Profile 변수에 넣는다


	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	driver.AddArgument("--start-maximized") ; 창 최대화 하기


	; ############ 여기서는 어차피 이전에 로그인 한 정보가 없으니 프로파일 읽는 코드를 넣을 필요가 없다 ############
	 

	; url 이 이래도 로그인 한 정보가 없으니 자동으로 로그인 화면으로 이동한다.
	url = https://vendoradmin.fashiongo.net/#/order/orders/new
	driver.Get(url)

;	MsgBox

	; 현재 화면이 로그인 화면이라면 로그인하기
	if(RegExMatch(driver.Url, "imU)fashiongo")){
		
		; driver 를 넘겨줘서 로그인 한 뒤 다시 driver 넘겨받기
		driver := FG_Login(driver)
		
		
		; 프로파일 경로 얻는 화면으로 이동
		URL = chrome://version/
		driver.Get(URL)

		; 프로파일 경로 찾아서 chrome_Profile 변수에 넣기
		Xpath = //*[@*='profile_path']
		chrome_Profile := driver.FindElementByXPath(Xpath).Attribute("innerText")

		; 프로파일 경로 확인
;		MsgBox, % chrome_Profile

		; 일단 창을 닫는다
		driver.close()	
		
	}





	; ########## 두번째 창 열기 ##########
	; 창을 연 뒤 첫번째 로그인 한 뒤 얻은 프로파일 경로를 읽는다
	; 그래도 로그인을 하라고 하기 때문에 또 로그인을 하게 되면 프로파일 경로는 첫 번째 프로파일 경로와 똑같고 맨 마지막에 \Default 이 붙는다.
	; 하지만 실제로 사용할 것은 첫 번째 프로파일 경로이다.

	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	driver.AddArgument("--start-maximized") ; 창 최대화 하기
	 
	 
	; 프로파일 경로 읽기
	driver.SetProfile(chrome_Profile) ; 이렇게 chrome_Profile 변수를 읽어서 처음 로그인 한 프로파일을 읽는다


	; 이렇게 새 창의 url을 넣어도 프로파일을 못 읽고 다시 로그인 하라고 한다
	url = https://vendoradmin.fashiongo.net/#/order/orders/new
	driver.Get(url)

;	MsgBox

	; 만약 로그인 화면이라면 로그인하기
	; if문을 썼지만 프로파일 못 읽고 로그인 화면으로 자동으로 빠진다.
	if(RegExMatch(driver.Url, "imU)fashiongo")){
		
		; driver 를 넘겨줘서 로그인 한 뒤 다시 driver 넘겨받기
		driver := FG_Login(driver)

		
		; 프로파일 경로 얻는 화면으로 이동
		URL = chrome://version/
		driver.Get(URL)

		; 프로파일 경로 찾아서 2NDchrome_Profile 변수에 넣기
		Xpath = //*[@*='profile_path']
		2NDchrome_Profile := driver.FindElementByXPath(Xpath).Attribute("innerText")

		; 프로파일 경로 확인
;		MsgBox, % 2NDchrome_Profile

		; 이렇게 두 번째 로그인 한 창을 닫아야 첫 번째 프로파일 저장한 프로파일 경로를 이용할 수 있다.
		driver.close()

		
		
		
	}

	; ########## 세번째 창 열기 ##########

	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	driver.AddArgument("--start-maximized") ; 창 최대화 하기
	 

	; 프로파일 경로 읽기. 두 번째 로그인 해서 프로파일 경로를 완성(?) 시켜 줬기 때문에 지금부터는 프로파일 경로를 제대로 읽는다
	driver.SetProfile(chrome_Profile) ; chrome_Profile 변수를 읽어서 처음 로그인 한 프로파일 경로를 읽는다


	; 이제는 로그인 화면으로 넘어가지 않고 제대로 뉴오더 화면으로 곧장 이동한다
	url = https://vendoradmin.fashiongo.net/#/order/orders/new
	driver.Get(url)

	MsgBox, OK 누르면 브라우저가 닫히면서 프로그램 종료합니다









Exitapp

Esc::
 Exitapp







	
	; 로그인 하기 
	FG_Login(driver){


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








