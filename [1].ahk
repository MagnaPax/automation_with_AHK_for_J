#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


/*
; goToURl_AfterLogIn_IfNeeded 메소드 테스트


#Include CNewBrow.AHK


	URL = https://vendoradmin.fashiongo.net/#/order/orders/new
	;~ URL = https://admin.lashowroom.com/orders_cur_month.php
	

	driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
	driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
	driver.AddArgument("--start-maximized") ; 창 최대화 하기	
	
	; 드라이버, URL 넘겨서 원하는 URL 로 이동한 뒤 드라이버 다시 받기
	driver := goToURl_AfterLogIn_IfNeeded(driver, URL)
	
	
MsgBox	
	
	driver.Get("http://google.com")

MsgBox

*/




	GetInfoFromLASPage(driver, PONumber)









Exitapp

Esc::
 Exitapp

