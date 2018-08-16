#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include FindTextFunctionONLY.ahk

/*
; 이거 원래 GetActiveBrowserURL.ahk 파일 안에 있던 함수인데 이게 메인에 선언되야 com 으로 처리한 변수들의 값이 유지되어 메인에서 사용할 수 있다.
Menu, Tray, Icon, % A_WinDir "\system32\netshell.dll", 86 ; Shows a world icon in the system tray

ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass"
*/


; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
BlockInput, Mouse





;~ LoginSkype()

LoginLAMBS()



;~ OpenLASHOWROOM()
;~ OpenFashionGo()



Run Outlook

URL = https://vendoradmin.fashiongo.net/#/order/orders/new
run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) URL   ;;이렇게 열린 크롬 창은 ChromeGet() 함수에 의해 재사용 될 수 있음 (새 창으로 열림)

;~ Run, C:\Program Files (x86)\UPS\WSTD\WorldShipTD.exe










MsgBox, 4100, , IT'S DONE























  OpenFashionGo(){

		
	; New Orders 검색화면 열기
	Loginname = customer3
	Password = Jo123456789
	URL = https://vendoradmin.fashiongo.net/#/order/orders/new
	WinMaximize, ahk_class IEFrame

	WB := ComObjCreate("InternetExplorer.Application")
	WB.Visible := true
	WB.Navigate(URL)
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		Sleep, 5000



	; 현재 url 얻는 부분
	nTime := A_TickCount
	sURL := GetActiveBrowserURL()
	WinGetClass, sClass, A
	If (sURL != ""){
		;MsgBox, % "The URL is """ sURL """`nEllapsed time: " (A_TickCount - nTime) " ms (" sClass ")"
	}
	Else If sClass In % ModernBrowsers "," LegacyBrowsers
		MsgBox, % "The URL couldn't be determined (" sClass ")"
	Else
		MsgBox, % "Not a browser or browser not supported (" sClass ")"


	; 얻은 현재 url이 로그인 화면이면 로그인 하기
	if(RegExMatch(sURL, "imU)login")){
			
		;~ wb.document.getElementById("tbUserID").value := Loginname  ;ID 입력
		;~ wb.document.getElementsByTagName("INPUT")[0].innerText := Loginname ;ID 입력
		
		;~ wb.document.getElementById("tbPassword").value := Password ; 비밀번호 입력
		;~ wb.document.getElementsByTagName("INPUT")[1].innerText := Password ;비밀번호 입력
			
			
		; 얘가 기계로 임력하는 걸 아는건지 위와 같은 일반적인 방법으로 로그인 하려고 하면 자꾸 에러가 나서 다음과 같이 이동 후 입력하는 방법 사용
		wb.document.getElementsByTagName("INPUT")[0].focus() ;ID 입력
		SendInput, % Loginname
		Sleep 100
			
		wb.document.getElementsByTagName("INPUT")[1].focus() ;비밀번호 입력
		SendInput, % Password
		Sleep 100
			
		wb.document.getElementsByTagName("BUTTON")[0].Click() ; 로그인 버튼 누르기
			
/*			
		; 로그인 후 New Orders 검색화면 열기
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100

		Sleep 3000
		wb.document.getElementsByTagName("I")[4].Click() ; 메뉴 바의 All Orders 의 ˅ 버튼 누르기
		wb.document.getElementsByTagName("A")[13].Click() ; All Orders 안에 있는 New Orders 버튼 누르기
			

		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			Sleep, 100
*/		
	}


	Sleep 3000

	IfWinExist, FashionGo Vendor Admin - Internet Explorer
		WinMinimize

   return
  }




OpenLASHOWROOM(){
 
		URL = https://admin.lashowroom.com/login.php
		

		Clipboard :=
		
		
		Loginname = jodifl
		Password = j123456789
		SecurityCode = 7864
		

		WB := ComObjCreate("InternetExplorer.Application")
		WB.Visible := true
		WB.Navigate(URL)
		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		   Sleep, 10



		wb.document.getElementById("uname").value := Loginname  ;ID 입력
		wb.document.getElementById("login_pwd").value := Password ; 비밀번호 입력
		wb.document.getElementsByTagName("INPUT")[2].Click() ; 로그인 버튼 누르기

		While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
		   Sleep, 10
		   
		;	MsgBox, found login
		
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 보안코드 입력하는 부분
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

		; 현재 url 얻는 부분
		nTime := A_TickCount
		sURL := GetActiveBrowserURL()
		WinGetClass, sClass, A
		If (sURL != ""){
			;MsgBox, % "The URL is """ sURL """`nEllapsed time: " (A_TickCount - nTime) " ms (" sClass ")"
		}
		Else If sClass In % ModernBrowsers "," LegacyBrowsers
			MsgBox, % "The URL couldn't be determined (" sClass ")"
		Else
			MsgBox, % "Not a browser or browser not supported (" sClass ")"



		; 얻은 현재 url이 Security Verification Code 입력 화면이면 보안코드 입력하기
		; https://admin.lashowroom.com/login_verify.php
		if(RegExMatch(sURL, "imU)login_verify")){
			wb.document.getElementById("verification_code").value := SecurityCode  ;Security Code 입력			
			wb.document.getElementsByTagName("INPUT")[2].Click() ; 로그인 버튼 누르기
			
			While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
			   Sleep, 10
		}
	

	Sleep 3000

	IfWinActive, LAShowroom.com Admin (JODIFL) -- Home Page - Internet Explorer
		WinMinimize	


	return
}




LoginLAMBS(){
	
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

	
	
	return
}




LoginSkype(){	

	Run, C:\Program Files (x86)\Skype\Phone\Skype.exe
	WinWaitActive, Skype

	; 커서 상태가 작업처리중이면 끝날때까지 기다리기
	while (A_cursor = "Wait")
		Sleep 3000

	IfWinActive, Skype
	{
		; 아이디 입력
		ControlSetText, Edit1, jodifl.han@outlook.com, Skype		
		Sleep, 3000


		; 커서 상태가 작업처리중이면 끝날때까지 기다리기
		while (A_cursor = "Wait")
			Sleep 3000


		; Next 버튼 찾아서 클릭
		Text:="|<Skype Next Button>*152$32.zzzzztwzzzyDDzzrVnzzxtQsNk6HAv9ranSOxtYk6DSQBzXrbXTuxtsngbSTC2Qlzzzzzs"
		;~ if ok:=FindText(569,539,150000,150000,0,0,Text)
		while ok:=FindText(569,539,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		}


		; 커서 상태가 작업처리중이면 끝날때까지 기다리기
		while (A_cursor = "Wait")
			Sleep 3000



		Sleep, 4000



		; 암호 입력 칸에 Passowrd라고 씌여있기 때문에 찾기
		Text:="|<Skype Password Button>*208$60.z000000003lU00000003lU00000003lbXluAFswzlYm3/AKAlXl0G23Cq4lXy3nXVSY4l3kAEktGY4l3k8EM9ma4lXkAoONnaAlXk7Hnkn3skzU"
		if ok:=FindText(747,487,150000,150000,0,0,Text)
		;~ while ok:=FindText(747,487,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  
		  Sleep, 5000
		  
		  SetKeyDelay 50,200 ; 사람처럼 키보드 천천히 입력하기 위해서
		  Send, Jo123456789
		  Send, {Enter}
		}

/*
		; Sign in 버튼 클릭하기
		Text:="|<Skype Sign in Button>*161$44.kFzzztztrzzzzzwzzzzzzzjvUM7tUMyna9yMn7gtbDaQwPCNntbDaraQyNnxgtbDaQCPCNntb8Cs6QyNnzztzzzzzzyTzzzzzkTzzzy"
		if ok:=FindText(956,539,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		}
*/
		
	}

	
	return
}








Exitapp

Esc::
 Exitapp
 Reload
 return