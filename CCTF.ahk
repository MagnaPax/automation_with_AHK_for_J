#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ ; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; 에러 메세지 경고창 안 뜨게 하는 함수
;~ ComObjError(false)


#Include %A_ScriptDir%\lib\

#Include FindTextFunctionONLY.ahk
#Include FG.ahk

#Include LAMBS.ahk
#Include CommLAMBS.ahk
#Include N41.ahk
#Include CommN41.ahk
#Include LAS.ahk

;~ #Include ChromeGet.ahk
#Include COM.ahk



global #ofCC_counter ; 램스에서 읽은 카드 갯수 저장하는 변수



L_driver := new LAMBS
N_driver := new N41
F_driver := New FG
LAS_driver := New LA


Arr_CSOS := object()
Arr_CC := object()
Arr_FGInfo := object()

/*
			URL = http://www.google.com
			driver := ChromeGet()
			driver.Get(URL)
			MsgBox, pause
*/

/*
; [##1##] CSOS 에서 정보 얻기
;~ Arr_CSOS := L_driver.getInfoFromCSOS()

; [##2##] 램스에서 카드 정보 읽어서 배열에 저장하기
;~ Arr_CC := L_driver.ReadingCCInfoFromLAMBS()

;~ MsgBox, % "value of variable #ofCC_counter : " #ofCC_counter
;~ MsgBox, % Arr_CC[1][3]
;~ MsgBox, % Arr_CC[1][5]


; [##3##] 램스에서 읽은 카드정보 N41에 입력하기
;~ WhereDoesThisComeFrom = 1
;~ ReadCCInfo_then_PutThatinN41CCWindow(Arr_CC, Arr_CSOS, N_driver, WhereDoesThisComeFrom)
*/






/*
driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
driver.AddArgument("disable-infobars") ; Close Message that 'Chrome is being controlled by automated test software'
driver.AddArgument("--start-maximized") ; Maximize Chrome Browser

driver.Get("https://vendoradmin.fashiongo.net/#/home")

driver.ExecuteScript("document.body.style.zoom = '100%';") ; Set the font of Chrome browser to 100%
driver.close() ; closing just one tab of the browser
*/











/*
;GUI Backgroud
Gui, Show, w250 h100, Pick Ticket Processing, AlwaysOnTop Window
WinSet, AlwaysOnTop, On, Pick Ticket Processing

;Input Customer PO Number
Gui, Add, Text, x22 y21 Cred , Customer PO #
Gui, Add, Edit, x102 y19 w100 h20 vCustomerPO,  ; MTR1DA6CA9799-BO1  ; OP122839178

;엔터 버튼
Gui, Add, Button, x22 y51 w200 h40 +default gClick_btn, Enter

;GUI시작 시 포커스를 CustomerPO 입력칸에 위치
GuiControl, Focus, CustomerPO


return
*/



;GUI Backgroud
Gui, Show, w350 h150, N41 Processing, AlwaysOnTop Window
WinSet, AlwaysOnTop, On, N41 Processing

;Input Start Order Id
Gui, Add, Text, x22 y21 Cred , Start SO #
Gui, Add, Edit, x92 y19 w100 h20 vStartSO#,  ;53493 ;49998 ;49993

;Input End Order Id
Gui, Add, Text, x22 y51 CBlue , End SO #
Gui, Add, Edit, x92 y49 w100 h20 vEndSO#,  ;11, 22, 33

;PO NO.
Gui, Add, Text, x22 y83 CGreen , Cust PO #`n(For Urgent Order)
Gui, Add, Edit, x92 y79 w100 h20 vCustomerPO,  ; MTR1D55EDC764

;ORDER ID
Gui, Add, Text, x22 y115 CBlack , Order ID Only`n(Skip Web Processing)
Gui, Add, Edit, x92 y109 w100 h20 vOrder_ID_Only, ; 54250



;엔터 버튼
Gui, Add, Button, x225 y19 w100 h110 +default gClick_btn, Enter



;GUI시작 시 포커스를 Cust PO # 입력칸에 위치
GuiControl, Focus, CustomerPO


return






Click_btn:

	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; To use the values which input on GUI
	Gui Submit, nohide
	GUI, Destroy


	WinClose, FashionGo Vendor Admin - Google Chrome
	WinClose, LAShowroom.com Admin (JODIFL) -- Orders Editing Page - Google Chrome
	WinClose, LAShowroom.com Admin (JODIFL) -- Orders PO Search Page - Google Chrome

	;~ StringReplace, CustomerPO, CustomerPO, %A_SPACE%, , All
	CustomerPO := Trim(CustomerPO)
	CustomerPO := RegExReplace(CustomerPO, "[^a-zA-Z0-9]", "")
	
	
	BuyerNotes := ""
	AdditionalInfo := ""
	StaffNotes := ""


; FG 오더 처리
if(RegExMatch(CustomerPO, "imU)MTR")){
	
	; FG 페이지에서 정보 읽어서 저장하기
	; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서
	; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
	Arr_FGInfo := F_driver.GettingInfoFromCurrentPage(CustomerPO)


	Arr_BillingAdd := Arr_FGInfo[1].Clone()
	Arr_ShippingAdd := Arr_FGInfo[2].Clone()
	Arr_CC := Arr_FGInfo[3].Clone()
	Arr_Memo := Arr_FGInfo[4].Clone()

	
	BuyerNotes := Arr_Memo[1]
	AdditionalInfo := Arr_Memo[2]
	StaffNotes := Arr_Memo[3]
	CC# := Arr_CC[2]


	; Order Status 를 Confirmed Orders로 바꾸기
;	F_driver.ChangeOrderStatusToConfirmedOrders()
	


	SoundPlay, %A_WinDir%\Media\Ring06.wav
	;~ MsgBox, 4100, Memo, %BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`nREADY TO UPDATE SHIPPING ADDRESS`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.
	MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE SHIPPING ADDRESS`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.
	;~ MsgBox, 4100, Memo, READY TO UPDATE SHIPPING ADDRESS`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.
	
	; No 눌렀으면 다시 시작
	IfMsgBox, No
	{
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
		Reload
	}
	
	; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
	N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
	
	

/*	; CC 입력할 때 활성화 하기

	; 입력할지 말지 결정하기 위해 카드 정보 입력창 열기
	N_driver.OpenRegisterCreditCard()

	Sleep 2000
	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 4100, Memo, CREDIT CARD NUMBER OF FG IS : `n%CC#%`n`n`nWOULD YOU LIKE TO TRANSFER CC INFO TO N41?`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.

	; Yes 눌렀으면 N41 에 카드 정보 입력하기
	IfMsgBox, Yes
	{
		; N41 에 카드 정보 입력하기
		N_driver.PutCCInfoInN41(Arr_CC, Arr_BillingAdd)
		N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
		
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
		Reload
	}
	
	; No 눌렀으면 CC 창 닫고 어플 다시 시작하기
	IfMsgBox, No
	{
		WinClose, Credit Card Management
		N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
		Reload
	}
*/

	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 262144, Title, Go to SO Manager Tab
	N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
	Reload
}






; LASHOWROOM 오더 처리
if(RegExMatch(CustomerPO, "imU)OP")){
	
	; LAS 페이지에서 정보 읽어서 저장하기
	; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서 
	; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
	Arr_FGInfo := LAS_driver.GetInfoFromLASPage(CustomerPO)


	Arr_BillingAdd := Arr_FGInfo[1].Clone()
	Arr_ShippingAdd := Arr_FGInfo[2].Clone()
	Arr_CC := Arr_FGInfo[3].Clone()
	Arr_Memo := Arr_FGInfo[4].Clone()

	
	BuyerNotes := Arr_Memo[1]
;	AdditionalInfo := Arr_Memo[2] ; 이 정보는 없음
;	StaffNotes := Arr_Memo[3] ; 이 정보는 없음
;	CC# := Arr_CC[2] ; 이 정보는 없음



	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`nREADY TO UPDATE SHIPPING ADDRESS`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.
	


	; No 눌렀으면 다시 시작
	IfMsgBox, No
		Reload
	
	; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
	N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
	
	; N41에 카드 정보가 있는지 확인하기 위해 카드 정보 입력창 열기
	N_driver.OpenRegisterCreditCard()

	Sleep 2000
	MsgBox, 262144, Memo, PLEASE CLICK Ok TO RESTART THE APPLICATION
	
	WinClose, Credit Card Management
	
	N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
	N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기

	Reload
}












; JODIFL WEB 오더 처리
if(RegExMatch(CustomerPO, "imU)JOD")){
	
	; Credit Sales Orders Small 탭에서 CustomerPO 검색한 뒤 열기
	L_driver.SearchPONumber(CustomerPO)
	
	; 램스에서 카드 정보 아닌 고객 메모 등 읽기
;	Arr_CSOS_Memo := L_driver.getInfoFromCSOS()
	

	; N41 창을 최소화하고 램스를 왼쪽모니터로 옮겨오기
	N41_wintitle := " N41"
	WinMinimize, %N41_wintitle%
	WinMove, LAMBS, , 0, 0
	WinMaximize, LAMBS

	
	; LAMBS 에서 카드 정보 읽어서 배열에 저장하기
	; 1~5 까지는 카드정보, 6~10까지는 주소정보 담겨있음
	Arr_CC_integration := L_driver.ReadingCCInfoFromLAMBS()
	
	
	; Arr_CC_integration에는 5개의 카드 값과 5개의 주소 값이 들어있지만 카드 갯수만큼만 루프 돌기
	;~ loop, 5{
	Loop, %#ofCC_counter%
	{		
		Arr_%A_Index%_CC := Arr_CC_integration[A_Index].Clone() ; 카드 정보는 배열의 1부터 넣으면 되지만
		Arr_%A_Index%_Billing := Arr_CC_integration[A_Index+5].Clone() ; 주소 정보는 6부터 있으니 인덱스에 5를 더한 6부터 시작한다
	}


	; cc정보가 없을수도 있지만 확인을 위해 변수에 값 넣기
	Name := Arr_1_CC[1]
	CC#_1 := Arr_1_CC[2]
	CC#_2 := Arr_2_CC[2]
	CC#_3 := Arr_3_CC[2]
	CC#_4 := Arr_4_CC[2]
	
;	MsgBox, % "Arr_1_CC[1]" . Arr_1_CC[1]
	
	CommN41.BasicN41Processing()
	
	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 4100, Memo, %BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`nREADY TO UPDATE SHIPPING ADDRESS`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.
	
	; No 눌렀으면 다시 시작
	IfMsgBox, No
		Reload


	; ## 고객 이름이 같은지 확인하기 ##	
	CommN41.ClickCustomerMasterTab()
		
	; To get Contact Name
	Text:="|<Contact>*149$41.S001005W00600O0QTSwRw1YmN9Bc2BYlm3E4P9YY6lAaHN9gwCAXTCC"
	if ok:=FindText(284,611,150000,150000,0,0,Text)
	{
		CoordMode, Mouse
		X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		MouseMove, (X+W)+80, Y+H//2
		Click
		Sleep 100
		Send, ^a^c
		Sleep 100
			
		;~ if Arr_CC[1] not contains %Clipboard%
		Name := % Arr_1_CC[1]
		; N41에 있는 이름이 웹이나 램스에서 가져온 값과 맞지 않으면 경고창 띄우기
		if Name not contains %Clipboard%
		{
			MsgBox, 4100, Alert, CONTACT NAME OF N41 IS NOT MATCHED WITH THE NAME ON FG`n`n`n`nWOULD YOU LIKE TO CONTINUE TO CHANGE ADDRESS INFO?`nIF YOU CLICK No, IT WILL RESTART THE APPLICATION.
				
			IfMsgBox, No
			{
				Reload
			}
			
			; Yes 눌렀으면 Contact에 값 입력하기
			IfMsgBox, Yes
			{
				MouseMove, (X+W)+80, Y+H//2
				Click				
				Send, % Name
			}
		}
			
	}
	
	

	; 입력할지 말지 결정하기 위해 카드 정보 입력창 열기
	N_driver.OpenRegisterCreditCard()
	
	MsgBox, 4100, Memo, CREDIT CARD NUMBER OF FG IS : `n%CC#_1%`n%CC#_2%`n%CC#_3%`n%CC#_4%`n`n`n`nWOULD YOU LIKE TO TRANSFER CC INFO TO N41?`nIF YOU CLICK No, THE APPLICATION WILL BE RESTART.

	; Yes 눌렀으면 N41 에 카드 정보 입력하기
	IfMsgBox, Yes
	{
		; 램스에 저장된 카드 갯수만큼만 루프 돌아서 카드정보 N41 에 입력하기
		Loop, %#ofCC_counter%
			N_driver.PutCCInfoInN41(Arr_%A_Index%_CC, Arr_%A_Index%_Billing)
		
		N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기		
		
		Reload
	}
	
	; No 눌렀으면 CC 창 닫고 어플 다시 시작하기
	IfMsgBox, No
	{
		WinClose, Credit Card Management
		
		N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
		N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
		
		Reload
	}

}

Exitapp








Esc::
Exitapp

^o::
URL = https://vendoradmin.fashiongo.net/#/home
CommWeb.OpenNewBrowser(URL)
return

^5::
MsgBox
N41_login_wintitle := "ahk_exe nvlt.exe"
WinWaitActive, N41_login_wintitle
WinMaximize, N41_login_wintitle
return

!z::
SendInput, ( available_allocate_qty > 0 )  
return

^r::
run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " )
return

F5::
Reload

GuiClose:
ExitApp 


















































F12::

	CN41_driver := New CommN41
	
	; Create Pick Ticket 버튼 클릭하기
	CN41_driver.ClickCreatePickTicketButton()

	WinMinimize, Pick Ticket Processing
	
	
	; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
	Sleep 3000	
	
	
	; Merge 확인 창
	IfWinActive, SO Manager
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		WinWaitActive, SO Manager
		IfWinActive, SO Manager
		{
			Sleep 500
			Send, {Enter}
			Sleep 1000			
			
			WinWaitActive, Pick Ticket ; Allocation 경고창
			IfWinActive, Pick Ticket
			{
				Sleep 500
				Send, {Left}
				Sleep 500
				Send, {Enter}
				Sleep 700
				
				
				WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
				IfWinActive, Pick Ticket
				{
					Sleep 500
					Send, {Enter}
					Sleep 800
					
					FromClickingPreAuthorizedButton_To_PrintOutPickTicket()
				}				
			}			
		}
	}


	WinWaitActive, Pick Ticket ; Allocation 경고창
	IfWinActive, Pick Ticket
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		
		WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
		IfWinActive, Pick Ticket
		{
			Sleep 500
			Send, {Enter}
			Sleep 800
			
			; Pick Ticket 창이 또 나오면 이전에 에러 메세지 창이 나왔을 것
			IfWinActive, Pick Ticket
			{
				MsgBox, MAYBE 'Warehouse is required!' ERROR HAS BEEN OCCURED`n`nWAREHOUSE INFO ON Sales Order OF THIS ORDER HAS TO BE MODIFIED.`n`nIF OK BUTTON ON Pick Ticket WINDOW ON N41, ALL INFO WILL BE SET AS DEFAULT.
				return
			}
			
			FromClickingPreAuthorizedButton_To_PrintOutPickTicket()
		}
	}
return	












; CBS 위한 처리
; pre authorized 버튼 누르지 않음
F11::

	CN41_driver := New CommN41
	
	; Create Pick Ticket 버튼 클릭하기
	CN41_driver.ClickCreatePickTicketButton()

	WinMinimize, Pick Ticket Processing
	
	
	; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
	Sleep 3000	
	
	
	; Merge 확인 창
	IfWinActive, SO Manager
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		WinWaitActive, SO Manager
		IfWinActive, SO Manager
		{
			Sleep 500
			Send, {Enter}
			Sleep 700			
			
			WinWaitActive, Pick Ticket ; Allocation 경고창
			IfWinActive, Pick Ticket
			{
				Sleep 500
				Send, {Left}
				Sleep 500
				Send, {Enter}
				Sleep 700
				
				
				WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
				IfWinActive, Pick Ticket
				{
					Sleep 500
					Send, {Enter}
					Sleep 800

					PrintOut()
				}				
			}			
		}
	}


	WinWaitActive, Pick Ticket ; Allocation 경고창
	IfWinActive, Pick Ticket
	{
		Sleep 500
		Send, {Left}
		Sleep 500
		Send, {Enter}
		Sleep 700
		
		
		WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
		IfWinActive, Pick Ticket
		{
			Sleep 500
			Send, {Enter}
			Sleep 800			
			
			; Pick Ticket 창이 또 나오면 이전에 에러 메세지 창이 나왔을 것
			IfWinActive, Pick Ticket
			{
				MsgBox, MAYBE 'Warehouse is required!' ERROR HAS BEEN OCCURED`n`nWAREHOUSE INFO ON Sales Order OF THIS ORDER HAS TO BE MODIFIED.`n`nIF OK BUTTON ON Pick Ticket WINDOW ON N41, ALL INFO WILL BE SET AS DEFAULT.
				return
			}			
			
			PrintOut()
		}
	}

return	

































^1::
;~ SetKeyDelay, 300
;~ SetKeyDelay 50,200
SetKeyDelay, 1000
;~ SetKeyDelay 300,200

Data = %Clipboard%

StringReplace, Data, Data, ', , All
StringReplace, Data, Data, -, , All
StringReplace, Data, Data, (, , All
StringReplace, Data, Data, ), , All
Data := Trim(Data)
StringUpper, Data, Data ; 대문자로 바꾸기

;~ StringLeft, Data, Data, 20  ; 왼쪽부터 20개 읽어서 저장하기

Send, %Data%
return




^2::
SetKeyDelay, 1000
Data = %Clipboard%

;~ RegExMatch(Data, "imU)(\d*)\.", SubPat)
;~ Data := SubPat1

Data := Trim(Data)
Send, %Data%
return




^3::
SetKeyDelay, 1000

Data = %Clipboard%

Data := RegExReplace(Data, "[^0-9]", "") ;숫자만 저장

StringReplace, Data, Data, ', , All
StringReplace, Data, Data, -, , All
StringReplace, Data, Data, (, , All
StringReplace, Data, Data, ), , All
StringReplace, Data, Data, %A_SPACE%, , All
StringReplace, Data, Data, `n, , All
StringReplace, Data, Data, `r, , All
StringUpper, Data, Data ; 대문자로 바꾸기
Data := Trim(Data)


;~ StringLeft, Data, Data, 20  ; 왼쪽부터 20개 읽어서 저장하기

Send, %Data%
return

























; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지 
FromClickingPreAuthorizedButton_To_PrintOutPickTicket(){
	
		CN41_driver := New CommN41
	
		; pre authorized 버튼 클릭
		Text:="|<pre-authorize Button>*205$16.001zzbzyTztzzY0SE1tzzbzyTztzzc01zzy"
		if ok:=FindText(718,129,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			
			Sleep 1000
			
			; Pre-Authorized 통과 됐거나 Declined 됐을 때
			WinWaitActive Credit Card Processing, , 4
			IfWinActive, Credit Card Processing
			{
				Sleep 500
				Send, {Enter}
				Sleep 500
				
				; Print 버튼 클릭
				Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
				if ok:=FindText(359,129,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, X+W//2, Y+H//2
					Click					
					
					Sleep 1000
					
					; 프린트 창 최대화 하기
					WinWaitActive, Pick Ticket Print
					WinMaximize, Pick Ticket Print
;~ /*					
					; 안에 있는 프린트 버튼 클릭
					Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
					if ok:=FindText(199,44,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, Y+H//2
						Click
						
						Sleep 500
						
						; 에러창 나오면 프로그램 다시 시작하기
						IfWinActive, Microsoft Visual C++ Runtime Library
						{						
							MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
							Reload
						}
						
						Send, {Down}
						Sleep 200
						Send, {Down}
						Sleep 200
						Send, {Enter} ; Print Now 눌러서 인쇄하기
						
						Sleep 3000
						WinActivate, Pick Ticket Print
						WinClose, Pick Ticket Print ; 프린트 창 닫기
						Sleep 700						
						
						;~ CommN41.runN41() ; N31 활성화 한 뒤 
						;~ CommN41.OpenSOManager() ; SO Manager 탭 열고 끝내기						
						;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
						
						; SO MANAGER 탭 누르고 끝내기						
						CN41_driver.ClickREfresh()
						
						Send, {Enter} ; 리프레쉬 버튼 누른 뒤 
						Sleep 700
						
						result := CN41_driver.DoesThisPickTicketApproved() ; Approved 됐는지 화면에서 찾아본 뒤 찾았으면 1을 리턴하고 못 찾았으면 0을 리턴
						if(result == 0){					
							MsgBox, Does This Pick Ticket Approved?							
						}						
						
						; SO Manager 탭 클릭해서 pick ticket 탭에서 나오기
						CN41_driver.OpenSOManager()
						
						; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
						; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
						CN41_driver.ClickREfreshButtonOnSOManager()
						

						Reload
						
					}
*/									
					
					
				}
							
				
			}
				
			; CC 가 없어서 업데이트 할거냐고 물을 때
			IfWinActive, Pick Ticket
			{
;				SoundPlay, %A_WinDir%\Media\Ring06.wav
;				MsgBox, 262144, Title, CC update`n`nCHECK THE SHIP VIA`n`nCLICK OK TO CONTINUE
				
				Sleep 300				

				Send, {Right}
				Sleep 200
				Send, {Enter}
				Sleep 500			


				
				; Print 버튼 클릭
				Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
				if ok:=FindText(359,129,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, X+W//2, Y+H//2
					Click					
					
					Sleep 1000
					
					; 프린트 창 최대화 하기
					WinWaitActive, Pick Ticket Print
					WinMaximize, Pick Ticket Print
;~ /*					
					; 안에 있는 프린트 버튼 클릭
					Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
					if ok:=FindText(199,44,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, Y+H//2
						Click
						
						Sleep 500
						
						; 에러창 나오면 프로그램 다시 시작하기
						IfWinActive, Microsoft Visual C++ Runtime Library
						{						
							MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
							Reload
						}
						
						Send, {Down}
						Sleep 200
						Send, {Down}
						Sleep 200
						Send, {Enter} ; Print Now 눌러서 인쇄하기
						
						Sleep 3000
						WinActivate, Pick Ticket Print
						WinClose, Pick Ticket Print ; 프린트 창 닫기
						Sleep 700						
						
						;~ CommN41.runN41() ; N31 활성화 한 뒤 
						;~ CommN41.OpenSOManager() ; SO Manager 탭 열고 끝내기						
						;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
						
						SoundPlay, %A_WinDir%\Media\Ring06.wav
						MsgBox, 262144, Title, NO CC INFO ON THIS CUSTOMER`n`n`n`nCHECK THE SHIP VIA`n`nCLICK OK TO CONTINUE

						; SO MANAGER 탭 누르고 끝내기
						CN41_driver := New CommN41
						CN41_driver.OpenSOManager()
						
						
						; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
						; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
						CN41_driver.ClickREfreshButtonOnSOManager()						
						
						Reload
						
					}
				}

				;~ Reload
				
			}

		}

	return
}








; 프린트 하기
PrintOut(){
	
	; Print 버튼 클릭
	Text:="|<Print Button>*165$17.0007s08A0E40U81TES0wZx982GTwY01802TzwzztU0lzz000E"
	if ok:=FindText(359,129,150000,150000,0,0,Text)
	{
		CoordMode, Mouse
		X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		MouseMove, X+W//2, Y+H//2
		Click					
					
		Sleep 1000
					
		; 프린트 창 최대화 하기
		WinWaitActive, Pick Ticket Print
		WinMaximize, Pick Ticket Print
				
		; 안에 있는 프린트 버튼 클릭
		Text:="|<Print Button2>*186$18.000TzyQ0yQ0SQ3CQ0CI0C40O002002002002002o0Dw0Dw0TzzzU"
		if ok:=FindText(199,44,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
						
			Sleep 500
					
			; 에러창 나오면 프로그램 다시 시작하기
			IfWinActive, Microsoft Visual C++ Runtime Library
			{						
				MsgBox, 262144, ALERT, RESTART THE APPLICATION DUE TO WARNING WINDOW`nYOU SHOULD KEEP CURRENT PICK TICKET NUMBER
				Reload
			}
						
			Send, {Down}
			Sleep 200
			Send, {Down}
			Sleep 200
			Send, {Enter} ; Print Now 눌러서 인쇄하기
						
			Sleep 3000
			WinActivate, Pick Ticket Print
			WinClose, Pick Ticket Print ; 프린트 창 닫기
			Sleep 700						
					
			;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, CBS_ORDER, CBS or CALL FOR CC
			
			; SO MANAGER 탭 누르고 끝내기
			CN41_driver := New CommN41
			CN41_driver.OpenSOManager()

			; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
			; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
			CN41_driver.ClickREfreshButtonOnSOManager()
			
			Reload
		}

	}	

	return	
}
