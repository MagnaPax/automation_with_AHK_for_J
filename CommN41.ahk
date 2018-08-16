#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;~ class CommN41{
class CommN41{

	; N41 로그인
	runN41(){
		
		Sleep 500
		
		N41_login_wintitle := " N41"
		N41_login_wintitle := "ahk_class FNWND3126"
		
		
;		WinWaitActive, %N41_login_wintitle%
;		WinMaximize, %N41_login_wintitle%

		WinMaximize, ahk_exe nvlt.exe


		; 이미 N41 열려있으면 메소드 중단하고 나오기
;		IfWinExist, %N41_login_wintitle%
		IfWinExist, ahk_exe nvlt.exe
			return


		Run, nvlt.exe, C:\NVLT
			
		; 로그인 창이 활성화 될 때까지 기다리기
		WinWaitActive, %N41_login_wintitle%
			
		; 아이디 입력
		ControlSend, Edit1, c123, %N41_login_wintitle%
		Sleep 200
			
		; Ok 버튼 클릭
		ControlClick, Button2, %N41_login_wintitle%, , l
			
		; 로그인 창이 닫힐때까지 기다리기
		WinWaitClose, %N41_login_wintitle%
		
		
		Sleep 500
		
		; 로그인 아이디 틀렸으면 나오는 경고창		
		IfWinActive, N41 Log In
			MsgBox, login again
		
		; 작업창 활성화 될때까지 기다리기
		WinWaitActive, %N41_login_wintitle%
		
		Sleep 4000
		
		while (A_cursor = "Wait")
			Sleep 3000
		
		Sleep 1000
		
		
		return
	}
	
	
	BasicN41Processing(){
		
		N41_login_wintitle := " N41"
		
		CommN41.runN41()
		
		IfWinActive, Connection Error
			MsgBox, 4100, Connection Error Warnning, Click Ok to continue
		
;		WinActivate, %N41_login_wintitle%
		WinActivate, ahk_exe nvlt.exe
		
	}
	














	; Customer Master Tab 클릭하기
	ClickCustomerMasterTab(){
		
		; Customer Master Tab 클릭하기
		Text:="|<Customer Master Tab>*170$81.Tzzzzzzzzzzzzw0D00000000000026000000000000kM000000000003y000000000011y8D00E0000Mk8n168020000361MEEUFSxnyQsRr42642+GPGIK2e0Eb0UFMG+GyULH1tg428mFGI42GUDMElHGHOGkUEI27z3vvniGHo22szzs00000000007zz00000000001zzs0000000000Dz000000000004"
		if ok:=FindText(609,104,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		else{
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, No CM Tab Warnning, Please Open Customer Master
		}			
		
		return		
	}
	
	
	; Customer Master 에서 카드 정보 창 아이콘 클릭해서 열기
	OpenRegisterCreditCard(){
		
		; N41 로그인
		CommN41.BasicN41Processing()

		; Customer Master Tab 클릭하기
		CommN41.ClickCustomerMasterTab()
		
		Sleep 500
		
		; 카드 아이콘 클릭
		Text:="|<CC ICON>*183$17.zzzzlzs1w03s07m03Y0300600CDzQTyw1Zw03s0Dzzw"
		if ok:=FindText(697,130,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click			
		}		
		
	}

	
	; SO Manager 클릭
	OpenSOManager(){		
		
		; N41 로그인
		CommN41.BasicN41Processing()
		
		; SO Manager 클릭
		Text:="|<SO Manager>*161$75.zzzzzzzzzzzzzz000000000000A000000000001E00000000003t00000000001z47bkMk00000CTV5X36000001U4888RrbrXnbDUVl12e2m2mWkz41c8LHoHoLo1wU512GWWWWUUVY8gMEIoIoK4CQVsy22yWySSVz400000000E03kU0000000w00040000000000zzU0000000004"
		if ok:=FindText(275,104,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		}
		
	}
	
	; CustomerMaster 에 있는 뉴 버튼 클릭하기
	ClickNewButtonOnCustomerMaster(){
	
		Sleep 700		

		Text:="|<New Button on Customer Master>*188$35.zzzzzzzzzzzzzzzzzzzzzzzxz00Tznw00zzztzxzzznzvzvzbzrzqzDzzz9yTzzzTwztzyztznzxzs61zzTzzDzwzzyTzzzzzzzzzzzzzzzzzzzjzzzzyTzzzzw"
		if ok:=FindText(239,130,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Sleep 500
		}

		return
	}


	; Sales Order 클릭
	ClickSalesOrderOnTheMenuBar(){	

		; Sales Order on the Menu bar
		Text:="|<Sales Order on the Menu bar>*137$73.zw0000000000E3000000000081E0000000004DY0040000102Tl1s2007k0U1CTV41006A0E0a0EUwbD22ttnns8Q1IIV1F558z41bfv0UcWyY7m0IJ0MEIFEG8N2+OkYAO9g9CQVtxDS3t3nobwE000000000Ew80000000008040000000007zy0000000002"
		if ok:=FindText(122,57,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}

		; 커서 상태가 작업처리중이면 끝날때까지 기다리기
;		while (A_cursor = "Wait")
;			Sleep 2000				

		
		return
	}


	; Add(+) Button
	ClickAdd(){		

		; Add(+) Button
		Text:="|<Add(+) Button of Sales Order>*147$14.0000030180G05UDz4ztTyDz0S07U1s0A000008"
		if ok:=FindText(458,128,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}

				
		; 커서 상태가 작업처리중이면 끝날때까지 기다리기
		while (A_cursor = "Wait")
			Sleep 2000				

		Sleep 1000
		
		return
	}


	; Save Button
	ClickSave(){

		; Save Button
		Text:="|<Save Button of Sales Order>*167$58.k0k0w7U0TXU70401011j0w0G0Y041S7U182E0E4ww04U901TFzU0G0Y0w13w0182E2LoDk04zt090FzU0E040bzDD013sE201sS04EF080D0w0Fz40zzs1k14QE3zw0007lT0A00000Tzw0TzU"
		if ok:=FindText(381,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
				
		
		return
	}
	
	
	
	; 맨 위 메뉴바에 있는 Customer 클릭하기
	ClickCustomerOnTheMenuBar(){

		Text:="|<Customer on the top menu>*144$86.000000000000000000000000000000000000000000007U0000000000Dk2600000000002010k0000000000U09k00000000008E3l1s020000002S0EFW00U000000bU00E8jStzCRs09U2242+GPGIKG02Q0C10WkYIZx600X7YEE8X959EEM081u26+OGPGK4G028VUkyywvYYx7U0bMTw00000000009rzz000000000027zzk0000000000Uzw00000000000Dk0000000000000000000000000000U"
		if ok:=FindText(1969,56,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		
		return
	}
	
	
	; Customer Master 탭에 있는 List 클릭하기
	ClickListOnCustomerMaster(){
		
		MouseClick, l, 261, 164		

		Text:="|<List on Customer Master>*184$59.s0TzU00001k0zz000003U10200000702To220U0C040840100Q09zE8/r00s0E0UEQY01k0bx0Ui803U102117E0702To23aU0C04087ptU0Q083k00000s0E7U00001k0UC000003U1zs000004"
		if ok:=FindText(243,155,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}	
		
		return
	}
	
	
	; Customer Master 탭에 있는 검색조건이 이메일로 된 공란 클릭
	ClickTheBlankModeByEmail(){
		
		MouseClick, l, 1030, 133
		
		return
	}
	
	
	; Customer Master 탭에 있는 Create SO 클릭하기
	ClickCreateSO(){		

		Text:="|<Create So>*183$16.zzzs0z03tzD4AtUHaTCMAtkTbryNwtXnbA6TkNznU7zzzs"
		if ok:=FindText(581,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}
		
		return
	}
	
	; Pick Ticket 탭에 있는 Refresh Button 클릭하기
	ClickREfresh(){

		Text:="|<REFRESH BUTTON>*172$59.McA00U0000lEM0300001WU00C000035000zU03k6+000zU08EAI600lU0UTMcA00VU102lE00U10205WU010204Dv50032008U6+303600G0AI603y00c0Mc003y01U2lE000s01ztWU001U000351U0200008"
		if ok:=FindText(235,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 200
		}

		return
		
	}
	
	; Approved 됐는지 화면에서 찾아본 뒤 찾았으면 1을 리턴하고 못 찾았으면 0을 리턴
	DoesThisPickTicketApproved(){

		Text:="|<CC Approved>*202$48.6000000360000003D7bbSnST96qqHHHHNYoonOznTYoonSknkqqoHAPHUrbYSAST0440000004400000U"
		if ok:=FindText(1637,299,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  return 1
		}
		
		return 0
	}
	
	
	ClickCreatePickTicketButton(){
		

		Text:="|<Create Pick Ticket Button>*157$16.00E0200F028zl428Ll1Tw5rELB1zw5zELx1047zkU"
		if ok:=FindText(1132,127,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}	
		
		return
	}
	
	
	; SO Manager 에 있는 refresh 버튼 클릭
	ClickREfreshButtonOnSOManager(){
		
		Sleep 500

		Text:="|<REFRESH BUTTON ON SO Manager>*186$34.800000U0z00207y00808Q00U60s020w1U087k600U60y020M1k081k600U3VU0CzzyzzvzjrzzzzVzzzzzzzzU"
		if ok:=FindText(308,128,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}
		
		return
	}
	
	
	; Customer Master 에 있는 리프레쉬 버튼 클릭하기
	ClickRefreshButtonOnCustomerMaster(){
		
		Text:="|<Refresh Button>*182$65.zzzzzzzzzzzU07zy1zzJ0z00Dzs1zw01yTzTztVztzzwzyzyTlzm00tzxzwTnzY01nzzzkzXz/zvbzzzny3yLzrDyTzXyDwjziTwzz7wzxTzS1UTz7bzyzyzznzz0Dzxzxzzbzz0Tzvzvzzzzzbzzk07zzzzzzzzzzw"
		if ok:=FindText(261,130,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}
				
		
		
	}
	
	
	; Sales Order 에 있는 리프레쉬 버튼 클릭하기
	ClickRefreshButtonOnSalesOrder(){		

		Text:="|<Refresh Button On Sales Order Tab>*169$34.00200000M00003U00U0Tk0100zU0801X0100260404080E0E0U1U1V00603600M07w01U0Ds0600700800M00001002"
		if ok:=FindText(260,129,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 300
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.ClickRefreshButtonOnSalesOrder()
			
		}
	}
	
	
	; 고객 코드 얻기 위해 Sales Order 에 있는 Customer 표시 찾아가서 엔터쳐서 클릭보드에 복사 후 CustCode 변수에 넣어서 리턴하기
	GetCustomerCode(){
		
		Clipboard := ""
		CustCode := ""
		
		Text:="|<Customer on Sales Order>*181$55.zzzzzzzzzk00000000M00000000A0000000060D00800030AU040001U42HrSTrbk61/999+OM30YsgobxA0UG7KOHUa0NdNd99HH07bbawYj9U00000000k00000000M00000000Dzzzzzzzzy"
		if ok:=FindText(247,219,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+7, Y+H//2
			Sleep 150
			
			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			CustCode := Clipboard
			Sleep 150

;			MsgBox, % CustCode
			
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
			if(!CustCode){
			;~ if(CustCode == ""){
;				MsgBox, no value in variable, restart the method
				CommN41.GetCustomerCode()
			}
			
			
			return CustCode
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.GetCustomerCode()
		}

	}
	
	
	; Customer PO 번호 얻어서 리턴하기
	GetCustPONumber(){
		
		Clipboard := ""

		Text:="|<Cust PO on Sales Order>*177$49.zzzzzzzzzzzzzzzzk0000000M0000000A00000006S00ES3kjN0088XAJc4bi4F2zo2GG28VBO19l1sEZh0YCUU8LynGnEE6NHDDDA8FsdU0000000k0000000M0000000Dzzzzzzzz"
		if ok:=FindText(573,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+5, Y+H//2
			Sleep 150		  
			
			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			CustomerPO := Clipboard
			Sleep 150
			
			return CustomerPO
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.GetCustPONumber()
		}				
		
		
	}
	
	
	
	; 왼쪽 메뉴바에 있는 Customer 클릭하기
	ClickCustomerMarkOnTheLeftBar(){


		Text:="|<>*140$101.00000000000000000000000000000000000U00000000000000020000000000000000Dzk00000000000000zzk6Ekk000000000148EAXXU00000000068oUR7D0000000000BEd0uSK0000000000O0G1xgA0000000000oCY3vwM0000000000gx86kkk0000000001zzUBVVU0000000001zy000000000000001kQ0000000000000010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000w00000000000000026000000000000000860000000000000009k000000000000000S83k0080000000000EEAk00k000000004000E6nnnnzXnrU000221UBanBqvAaBU0005s30PD6lgqzgS0007YE30q7BXNBUMC000DkE6NgqNgmNglg000VUk7ntwttavtXs0033zU00000000000007zz00000000000000Tzy00000000000000zzw000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000E0000000000000009U000000000000000G0000000000000000A0000000000000001w03l0A00000000000005a0M00000000000008Tank000000000FU0QPBgU000000000X00QnnzU000000000700BbbU0000000004C06P6Bg0000000000A07bANk000000000UQ000M00000000001Vs003U000000000040c00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000004"

		if ok:=FindText(55,798,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 150
		}
	}
	
	; 왼쪽 메뉴바에 있는 SO Manager 클릭하기
	ClickSOManagerOnTheLeftBar(){

		Text:="|<SO Manager>*146$58.SD1X000002966A00000888RnbXXnbsUVJ1F1FFEO25Yx4x5x0c8GIIIII4WFV1FFFFEHks45x5wwx00000000E00000000S02"
		if ok:=FindText(45,154,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Sleep 150
		}

	}
	
	
	; SO Manager 에 있는 Customer 표시 찾기. 고객 코드 입력하기 위함
	FindCustomerMarkToFillInTheBlank(){
		
		Text:="|<Customer on the top menu bar of SO Manager>*147$52.U000000020000000081s020000UMU08000210Wxvbwts42+GPGIKUE8g959TG10WAYIZ186+OGPGK4UDjjCt9DG00000000800000000U00000002"
		if ok:=FindText(497,127,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+20, Y+H//2
			Click
			Sleep 150
		}
		else
		{
			; 못 찾으면 계속 재귀호출 해서 찾아보기
			Sleep 700
			CommN41.FindCustomerMarkToFillInTheBlank()
		}
	}

	
	; Pick Ticket 탭에 있는 House Memo 에 메모 넣기 위해
	PutMemoIntoHouseMemoOnPickTicket(){
		
		Clipboard := ""
		
		WinActivate, ahk_class FNWND3126

		Text:="|<House Memo ON Pick Ticket>*161$37.000003cU0001oE0000u9t9ssRwYYYaCWHGQT7F9d3c3cYYgoloHnnlks00000Q00000C000007000003U00001k00000sn0001QNU003yAlnywr6NB9GTWoyYdVlOEGIksh9d+EQKXYZsC000007000003U00001s"
		if ok:=FindText(1099,226,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, (X+W)+2, Y+H//2
		  Click
		  Sleep 150
		  Send, {End}
		  Sleep 150
		  Send, {Enter}
		  Sleep 150
		}	
		
	}


	; Open Allocation 에 있는 Create Pick Ticket 버튼 클릭하기
	ClickCreatePickTicketButtonOnOpenAllocation(){
		
;		MsgBox, 262144, Title, Open Allocation 에 있는 Create Pick Ticket 버튼 클릭하기


		Text:="|<Create Pick Ticket Button on Open Allocation Of SO Manager>*177$248.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzk0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000800000000000000000000000400000000000004005000000Dk00Dk000000Ds0002000000000000010DzM000007z007z000000A1U0014000001k000000E211000001yE01yE03zy0404000WDU000AQ00180040zxU00000Ty00Ty00jyE1Dt01zl6A0007j000k001081M033007zU07zU0BzA0U080E8VXzTz1vrXzjjTUE20+00NU05zu05zu03jb08zW05wEkSoykSxhUPPTM40U0U03k0100U100U0xnk200U1zwA77zAAznkyrba1080800M00Tzs0Tzs0D9w0XsE0TT1Xlkn3zwwRhttUE20200D007zy07zy03Yj040407nkMwyAknzhaPzSM40U0U06M01zzU1zzU0nxk10601zw3vxzAMTTTySza1080803300Tzs0Tzs09zg0Ay00Tz00k00000000000E3zy00000402040200zx018007zk0C000000000004000000000Ty00Ty00TzU0I00104000000000000000000000003z003z0000006000Tz2"
		if ok:=FindText(1185,332,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)-10, Y+H//2
			Click			
			Sleep 150
		}		
	}
	
	
	; Sales Order 탭에 있는 Memo
	; 메모값 읽기 위해
	MemoOnSalesOrderTab(){
		
		Clipboard := ""

		Text:="|<Customer Memo On Sales Order Screen>*191$45.zzzzzzzw0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000w0000007U06M000w00n0007U06NtzSw00z9heTU07vx9nw00hM9CTU05dh9Hw00hD9/rU000000w0000007U000000w0000007U000000w0000007U000006w000001zU000006w000000zU000000w0000007U000000w0000007U000000w0000007U000000w0000007U000000zzzzzzzzU"

		if ok:=FindText(1030,249,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, (X+W)+2, Y+H//2
		  Click
		  Sleep 150
		  
		  Send, ^c
		  Sleep 150
		  
;		  MsgBox, % Clipboard




		; 클립보드 내용 CustMemoOnSOTab 변수에 넣기
		Sleep 700
		CustMemoOnSOTab := Clipboard	


		; web 주문에 메모가 있다면 추출한 뒤 변수에 저장함
		CustMemoOnSOTab := RegExReplace(CustMemoOnSOTab, "(.*)CreditCard\sID\sfor\spayment:\s.*", "$1")  ; $1 역참조를 사용하여 메모 내용을 돌려준다

		; FG 주문에 메모가 있다면 추출한 뒤 변수에 저장함
		CustMemoOnSOTab := RegExReplace(CustMemoOnSOTab, ".*notes:(.*),\sOrder.*", "$1")  ; $1 역참조를 사용하여 메모 내용을 돌려준다

		; LAS 주문에 메모가 있다면 추출한 뒤 변수에 저장함
		CustMemoOnSOTab := RegExReplace(CustMemoOnSOTab, ".*LA,\s(.*).*", "$1")  ; $1 역참조를 사용하여 메모 내용을 돌려준다

		; LAS 주문의 메모를 저장한 변수에 None 이 있다면 메모가 없다는 뜻이니까 그냥 지운다
		if(CustMemoOnSOTab == "None"){
			CustMemoOnSOTab := ""
		}
		

		StringUpper, CustMemoOnSOTab, CustMemoOnSOTab ; 대문자로 바꾸기
		  
		return CustMemoOnSOTab
		  
		}

	} ; MemoOnSalesOrderTab() 메소드 끝


	
	; Sales Order 탭에 있는 House Memo 읽어서 리턴
	HouseMemoOnSalesOrderTab(){
		
		Clipboard := ""
		

		Text:="|<house memo>*140$32.W00008U00028t9ssyGGGG8YYb7m998N0WGGHG8X7bXU00000000000000000000000000000000240000n0000AlnyQ3AYYd0VDd+E/G2GY2oYYd0Z799c"
		if ok:=FindText(1033,289,150000,150000,0,0,Text)
		{
			
			Clipboard := ""
			
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+5, Y+H//2
			Click
			Sleep 150
			  
			Send, ^c
			Sleep 150
			
			
	;		MsgBox, % "house memo`n`n" . Clipboard



			; 클립보드 내용 CustMemoOnSOTab 변수에 넣기
			Sleep 700
			HouseMemoOnSOTab := Clipboard	


			HouseMemoOnSOTab := RegExReplace(HouseMemoOnSOTab, "Staff only notes:(.*)", "$1")  ; $1 역참조를 사용하여 Staff only notes: 이외의 메모 내용이 있으면 변수에 저장
			
			StringUpper, HouseMemoOnSOTab, HouseMemoOnSOTab ; 대문자로 바꾸기

			  
			return HouseMemoOnSOTab
			  
			}

	} ; HouseMemoOnSalesOrderTab() 메소드 끝
	



	
	; 화면의 4분변에 있는 Opne Allocation 의 빈 공간 클릭하기
	; 정확히 클릭하기 위해 왼쪽 메뉴바에 있는 Customer 찾아서 거기서 오른쪽으로 이동해서 클릭하기
	ClickEmptySpaceOnOpenAllocationArea(){

		Text:="|<>*140$101.00000000000000000000000000000000000U00000000000000020000000000000000Dzk00000000000000zzk6Ekk000000000148EAXXU00000000068oUR7D0000000000BEd0uSK0000000000O0G1xgA0000000000oCY3vwM0000000000gx86kkk0000000001zzUBVVU0000000001zy000000000000001kQ0000000000000010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000w00000000000000026000000000000000860000000000000009k000000000000000S83k0080000000000EEAk00k000000004000E6nnnnzXnrU000221UBanBqvAaBU0005s30PD6lgqzgS0007YE30q7BXNBUMC000DkE6NgqNgmNglg000VUk7ntwttavtXs0033zU00000000000007zz00000000000000Tzy00000000000000zzw000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000Tzzzzzzzzzzzzzzzz000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000E0000000000000009U000000000000000G0000000000000000A0000000000000001w03l0A00000000000005a0M00000000000008Tank000000000FU0QPBgU000000000X00QnnzU000000000700BbbU0000000004C06P6Bg0000000000A07bANk000000000UQ000M00000000001Vs003U000000000040c00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000004"

		if ok:=FindText(55,798,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, (X+W)+200, Y+H//2
		  Click
		  Sleep 150
		}
	}
		
	
	
	; Open Allocation 의 Chk 전체 선택하게 하기
	Click_Chk_On_OpenAllocation(){
		
		

		Text:="|<Chk on Open Allocation of SO Manager>*152$17.zzy00000000000000000000000DU0lU11T22a4548+8MokTD00E00U00000000000000001zzw007zzk00000000000003m2AY4ED/UGR0Ym19qOGbYYk"
		if ok:=FindText(953,342,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)-5
			Sleep 150
			Click, 2
			MouseMove, A_ScreenWidth / 2, A_ScreenHeight / 2
			
			Sleep 150
		}

	}
	

	; Open Allocation 화면의 체크박스가 체크됐을때
	Che_is_Checked(){
		

		Text:="|<Checked>*137$9.0830oAn3kA4"
		if ok:=FindText(953,377,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Sleep 300
			return 1
		}

/*
		Text:="|<Chk on Open Allocation of SO Manager is checked>*145$20.U008002D88aG290wiE9+Y2H90YuNd+XmGM002000U008002000U008002000zzzzk0Tw07z05zk3Tw1bzElzqMTww7z61zk0Tw07zzzs"
		if ok:=FindText(1065,368,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  ;~ MouseMove, X+W//2, Y+H//2
		  Sleep 300
		  ;~ MsgBox, the che 체크됐음 1 리턴함 - Che_is_Checked()
		  return 1
		}
*/

		return 0		
	}
	
	
	; SO Manager 화면에서 Last Inv. Dt 날짜 찾아서 리턴하기
	getLastInvDateOnSOManager(){
		
		Clipboard := ""

		Text:="|<Last Inv. Date>*127$80.k000A0k0001zVg00030A0000MAP0000k3000061ak3sSS0nwUk1UPw1XAn0AtgM0M6P0Mm0k3AP6061ak0wsA0n6NU1UNg1v7n0Alak0M6P0MkCk3ANg061ak4A1g0n6C01UNg1b8n0AlXU0MAPzDHss3AMMk7y7U"
		if ok:=FindText(587,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H//2)+30
			Sleep 150
/*			
			Sleep 2500			
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 1000
*/

			; 날짜 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			lastInvDate := Clipboard
			Sleep 150
			
			return lastInvDate
		}
	}
	
	
	; SO Manager 화면에서 priority 번호 읽어서 리턴하기
	getPriorityOnSOManager(){
		
		Clipboard := ""

		Text:="|<Priority On SO Manager>*126$27.0k00M000000007qDXvsn6QS6Mn3km3MS6EP3km3MS6Mn3kn6MS6DX3U"
		if ok:=FindText(359,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H//2)+30
			Sleep 500
			;~ Sleep 2000
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 1000
			
			; Priority 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			priority# := Clipboard
			Sleep 150
			
			return priority#
			
		}
	}
	
	; SO Manager 에 있는 Open SO 에 보낼 아이템 있는지 확인하기
	checkOpenSoIfThereAreItemsShipOut(){
		
		Clipboard := ""

		Text:="|<Color on Open SO of SO Manager>*109$46.3s00M000zk01U0073U06000M700M00308DVVw5w01z6DsTk0CCNllX00kNa36A031aMAEk0A6NUl30AkNa3461n1aMAEQCCCNll0zkTlXy40y0y67kEU"
		if ok:=FindText(1364,159,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+20
			Sleep 150
			
			; 마우스 오른쪽 버튼 클릭 후 Filter 메뉴 위에서 엔터치기
			Send, {RButton}			
			Loop, 10
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			; 필터 입력창 나올때까지 기다리기
			WinWait, Filter and Sort - SO Manager : d_so_manager_sod
			Sleep 200
			
			; 필터 입력창에서 조건 입력하기
			Text:="|<Apply Filter of Filter and Sort Window>*147$81.000000000000A0k00A0DhY0003UC001U1UBU000S9syTDAABySS03lD6PBtVthaH00nPAnNhsABhzM06PTaPBj1Vhg300zSAnNgkABgqM0CRkrnta1VhnX01Vc0kM0k000000000630Q000000004"
			if ok:=FindText(380,551,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, (Y+H//2)+50
				Sleep 100
				Send, {LButton}
				SendInput, ( available_allocate_qty > 0 )
				Sleep 100
			  
				; ok 버튼 클릭하기
				Text:="|<OK on Open Allocation>*177$67.zzzzzzzzzzzk0000000000M0000000000A000000000060000000000300000000001U0000000000k0003snU000M0003aPU000A0001XjU00060001krk00030000sPs0001U000QBw0000k0006Cr0000M0003iNk000A0000yAs00060000000000300000000001U0000000000k0000000000M0000000000A0000000000600000000003zzzzzzzzzzzU"
				if ok:=FindText(1018,616,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, X+W//2, Y+H//2
					Click
					Sleep 150
				  
				  
					Clipboard := ""
					isThereItemsOnOpenSo = 1

					Text:="|<Color on Open SO of SO Manager>*109$46.3s00M000zk01U0073U06000M700M00308DVVw5w01z6DsTk0CCNllX00kNa36A031aMAEk0A6NUl30AkNa3461n1aMAEQCCCNll0zkTlXy40y0y67kEU"
					if ok:=FindText(1364,159,150000,150000,0,0,Text)
					{
						CoordMode, Mouse
						X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
						MouseMove, X+W//2, (Y+H)+20
						Sleep 150
						
						; 마우스 오른쪽 버튼 클릭 후 Filter 메뉴 위에서 엔터치기
						Send, {RButton}			
						Loop, 4
						{
							Sleep 100
							Send, {Down}
							;~ Sleep 150
						}
						Send, {Enter}
						Sleep 150
						
						isThereItemsOnOpenSo := Clipboard
						Sleep 150
						
						; 화면에 아이템이 없으면 1 리턴하기
						if(!isThereItemsOnOpenSo){
							return 1
						}
					}
				  
				  
				  
				}
			}

		}
		
		return 0

	} ; checkOpenSoIfThereAreItemsShipOut() 함수 끝
	
	
	
	; SO Manager 의 3사분변의 Pick Ticket 에 오픈된 주문이 있는지 확인 한 뒤 있으면 1 리턴하기
	checkPickTicketSectionToFindIfPendingOrderExists(){
		
		;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
		;~ MsgBox, 262144, Title, 펜딩 오더가 있는지 확인합니다
		
		Sleep 500
		Clipboard := ""
		pickDate = 1

		Text:="|<Pick Date on Pick Ticket of SO Manager>*106$81.zkk0A00Ts00007z601U03zU01U0UQ00A00EC00A041U01U020k01U0UAkwAA0E33sS7o1aDlX020MzXlzUQnaAk0E3ACAQTz6MNg020M0lX1zkn0DU0E30SATw06M1w020MzlXzU0n0Ak0E3D6AM406MNX020lUlX0U0nbAM0ECASAQA06DlVU3zVzllzU0kwAA0Ts7q77o"
		if ok:=FindText(1390,518,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+23
			
			; Pick Date 밑에서 마우스 오른쪽 버튼 클릭 후 Pick Date 위에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			pickDate := Clipboard
			Sleep 150
			
;			MsgBox, % pickDate
			
			; pickDate 변수에 값이 없으면(펜딩 오더가 없으면) 1 리턴하기
			if(!pickDate){
;				MsgBox, 펜딩 오더가 없으니 1 리턴하기
				return 1 
			}			
			
		}
		
		; 펜딩 오더가 있으면 0 리턴
		return 0
		
		
	} ; checkPickTicketSectionToFindIfPendingOrderExists() 메소드 끝
	
	
	; SO Manager 에서 카드가 있는지 없는지 확인 후 있으면 1 리턴하기
	checkCC(){
		
		Clipboard := ""
		#ofCC = 0
		

		Text:="|<CC on Customer List of SO Manager>*127$22.D0DUk3X20A6y1U0U6020M081U0U6020M780kMU3X203sU"
		if ok:=FindText(524,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+20
			
			; CC 밑에서 마우스 오른쪽 버튼 클릭 후 Copy 위에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 100
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			#ofCC := Clipboard
			Sleep 150
			
;			MsgBox, 262144, Title, #ofCC : %#ofCC%
			
			; #ofCC 변수에 값이 있으면 참 값인 1 리턴하기
			if(#ofCC){
				return 1
			}
			
			
			
		}
		
		return 0
		
	} ; checkCC() 메소드 끝
	
	
	
	;Sales Order 에서 Order Type 값 얻기
	getOrderType(){
		
		Clipboard := ""
		orderType := ""
		


		Text:="|<Order Type on Sales Order>*166$53.S0600z001a0A00M0024vtnUq/lo94oo1YoosGNjc3B9jkYnEE6CHEn8aaUAMaawFwt0Mlss000001200000006408"
		if ok:=FindText(735,187,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+30, Y+H//2
			Sleep 150
			
			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			orderType := Clipboard
			Sleep 150

;			MsgBox, % CustCode
			
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
			if(!orderType){
			;~ if(CustCode == ""){
;				MsgBox, no value in variable, restart the method
				CommN41.getOrderType()
			}
			
			
			return orderType
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.getOrderType()
		}		
		
	} ; getOrderType 메소드 끝
	
	
	
	; SO Manager 화면의 왼쪽 밑 Pick Ticket 섹션에서 Pick Date 날짜 가져오기
	getPickDateOnPickTicketSectionOfSOManager(){
		
		i = 1
		Clipboard := ""
		pickDate := ""
				

		Text:="|<PickDate On PickTicket Section Of SO Manager>*122$83.zzzzzzzzzzzzzy000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000007z300k03zU0000Dz601U07zU01U0M700300A3U0300k600600M300601UAkwAA0k33sS7n0NXwMk1U6DswTq1nCMn030Aksllzz6MNg060M0lX1zwAk3s0A0kDX7zk0NU7k0M1Xz6DzU0n0Ak0k3D6AM301a6Nk1UAMQMk603CQlU30slslkw06DlVU7zVzltzM0AD330Dy1xXlwU"
		if ok:=FindText(503,697,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, (Y+H)+25
			Sleep 150
			
			; 고객 코드명 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			pickDate := Clipboard
			Sleep 150

;			MsgBox, % pickDate
/*			
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
			if(!pickDate){
				
				if(i == 3){
					CommN41.getPickDateOnPickTicketSectionOfSOManager()
					i++					
					
				}

			}
*/			
			
			return pickDate
		}
		else
		{
			; 못 찾았으면 재귀호출해서 계속 찾기
			Sleep 500
			CommN41.getPickDateOnPickTicketSectionOfSOManager()
		}						
		
		
		
		
	} ; getPickDateOnPickTicketSectionOfSOManager 메소드 끝
	
	
	
	; Sales Order 탭에서 고객의 할인율 찾아서 할인받았으면 1 리턴하기
	FindCustDCRate(){
		
		Clipboard := ""

		Text:="|<Cust DC Rate>*177$67.zzzzzzzzzzzk0000000000M0000000000A00000000006D008D3kT083AU044H88U41Y2Hr2B04PrCm19916U2999t0YsUXE1wQbwUG7EFc0YGG6NdNc8aMF99D7bba7Vs8rqRU0000000000k0000000000Tzzzzzzzzzzw"
		if ok:=FindText(409,201,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			;~ MouseMove, X+W//2, Y+H//2
			MouseMove, (X+W)+10, Y+H//2
			Sleep 150
			
			
			; 고객 코드명 위에서 할인율 얻기 위해 마우스 오른쪽 버튼 클릭 후 밑으로 4칸 내려서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			discountRate := Clipboard
			Sleep 150

			; 할인 받았으면 1 리턴하기
			if(discountRate){
;				MsgBox, % discountRate
				return 1
			}
			; 할인 안 받았으면 0 리턴하기
			else
				return 0			
			
		}

		
		
	} ; FindCustDCRate 메소드 끝
	
	
	
	
	
	; pick ticket 화면에서 정보 읽은 뒤 값들 리턴하기
	getInfoOnPickTicket_Then_ReturnThem(){
		
		Sleep 2000
		
		
		WinActivate, ahk_class FNWND3126		
		
		Clipboard := ""
		
		
		; pick # 찾기
		Text:="|<pick# on Pick Ticket>*174$30.000070000700007wU81TW081LWb/7zWdi3LwcA2rUcC7zUce2bUb92bU"
		if ok:=FindText(276,154,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, (X+W)+5, Y+H//2
			Sleep 150
			
			; pick #  위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
			
			pick# := Clipboard
			Sleep 150

;			MsgBox, % pick#

;~ /*
			; 만약 변수에 값이 안 들어갔으면 재귀호출해서 다시 처음부터 시작하기
;			if(!pick#){
;				MsgBox, 262144, Title, pick# 변수에 값이 없음. 다시 시작
;				CommN41.getInfoOnPickTicket_Then_ReturnThem()
;			}			
*/			
			
			loop{ ; 첫번째
				
				
				Clipboard := ""
				
				; Customer Code 얻기
				Text:="|<Customer on Pick Ticket>*174$48.zzzzzzzz000000070000000700000007S00E0007m00E0007UGSvnyQzUGGGGGabUGQGOHybUG7GOHUbnGnGGGabSSSPmGQb000000070000000700000007zzzzzzzzU"
				if ok:=FindText(267,203,150000,150000,0,0,Text)
				{
					CoordMode, Mouse
					X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
					MouseMove, (X+W)+5, Y+H//2
					Sleep 150
					
					; pick #  위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
					Send, {RButton}
					Loop, 4
					{
						Sleep 150
						Send, {Down}
						;~ Sleep 150
					}
					Send, {Enter}
					Sleep 150
					
					CustCode := Clipboard
					Sleep 150

;					MsgBox, % CustCode

;~ /*
					; 만약 변수에 값이 안 들어갔으면 루프 처음으로 돌아가서 다시 시작하기
;					if(!CustCode){
;						MsgBox, 262144, Title, CustCode 변수에 값이 없음. 다시 시작
;						CommN41.getInfoOnPickTicket_Then_ReturnThem()
;					}
*/					
					
					loop{ ; 두 번째
						
						Clipboard := ""
							
						; 작성한 날짜, 시간 얻기
						Text:="|<Update Date on Pick Ticket>*163$45.W033k10QE0MF083WST2BvbQGOMFd9DWHn2AtDwGSMFd93aHH2999vXnsSDgv0E00000M2000003U"
						if ok:=FindText(1658,200,150000,150000,0,0,Text)
						{
							CoordMode, Mouse
							X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
							MouseMove, (X+W)+5, Y+H//2
							Sleep 150
							
							; pick #  위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
							Send, {RButton}
							Loop, 4
							{
								Sleep 150
								Send, {Down}
								;~ Sleep 150
							}
							Send, {Enter}
							Sleep 150
							
							updDate := Clipboard
							Sleep 150

;							MsgBox, % updDate

;~ /*
							; 만약 변수에 값이 안 들어갔으면 루프 처음으로 돌아가서 다시 시작하기
;							if(!updDate){
;								MsgBox, 262144, Title, updDate 변수에 값이 없음. 다시 시작
;								CommN41.getInfoOnPickTicket_Then_ReturnThem()								
;							}
*/							
							
							; 변수들 중 한 개에라도 값이 아무것도 없으면 다시 시작
;							if(!pick# && !CustCode && !updDate){
;							MsgBox, 262144, Title, 변수들 중 값이 없는 변수가 있음. 다시 시작
;							CommN41.getInfoOnPickTicket_Then_ReturnThem()							
;							}
							
;		MsgBox, 262144, Title, 찾은값들 리턴하기 전에 확인해보기`n`n`n%pick#%`n`n`n%custCode%`n`n`n%updDate%
							
							return [pick#, CustCode, updDate]
							
							
							
						} ; if ends - 작성한 날짜, 시간 얻기						
				
						; 만약 그림 못찾았으면 재귀호출해서 다시 처음부터 시작하기
						;~ else if(!updDate){
;						else{
;							MsgBox, 262144, Title, updDate 그림 못 찾았음
;							CommN41.getInfoOnPickTicket_Then_ReturnThem()
;						}
						
						
					} ; loop ends - 두 번째

				} ; if ends - Customer Code 얻기
		
				; 만약 그림 못찾았으면 재귀호출해서 다시 처음부터 시작하기
				;~ else if(!CustCode){
;				else{
;					MsgBox, 262144, Title, CustCode 그림 못 찾았음
;					CommN41.getInfoOnPickTicket_Then_ReturnThem()
;				}			
							
				
				

			} ; loop ends - 첫번째

		} ; if ends - pick # 찾기
		
		; 만약 그림 못찾았으면 재귀호출해서 다시 처음부터 시작하기		
		;~ else if(!pick#){
		else{
			MsgBox, 262144, Title, pick# 그림 못 찾았음
			CommN41.getInfoOnPickTicket_Then_ReturnThem()
		}			
			
		
		; 중간에 리턴되지 않고 이 코드가 실행되면 뭔가 이상한것. 무한반복 되려나?
		CommN41.getInfoOnPickTicket_Then_ReturnThem()
		
	} ; getInfoOnPickTicket_Then_ReturnThem 메소드 끝
	


	; SO Manager 화면에서 주소를 읽어서 리턴
	getADDr(){
		
		Clipboard := ""
		addr := ""
		
		WinActivate, ahk_class FNWND3126
		

		Text:="|<Address 1 on SO Manager>*117$71.3U1UA00000027030M000000AP060k000000sq3wTbnsS7U3l4AtbCAN6FU1aAEm6MMm0U03ANVgAkUa1U06zv3MNXz7Vs0BUq6kn201UM0O1YAVa6410E0w1gNXAAN6FU1s3DlyMDXsy03U"
		if ok:=FindText(411,155,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, (Y+H)+20
			Sleep 150
							
			; 주소 위에서 마우스 오른쪽 버튼 클릭 후 코드 복사메뉴에서 엔터치기
			Send, {RButton}
			Loop, 4
			{
				Sleep 150
				Send, {Down}
				;~ Sleep 150
			}
			Send, {Enter}
			Sleep 150
							
			addr := Clipboard
			Sleep 150
			
			; 주소 못 읽었으면 재귀호출
;			if(!addr){
;				CommN41.getADDr()
;			}
			
			return addr
			
		}
				
		
		
		
	} ; getADDr() 메소드 끝
	
	
	
	



	
	
	
	
	
	
	
	
	
} ; 전체 클래스 class CommN41 끝
