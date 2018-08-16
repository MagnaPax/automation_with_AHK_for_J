#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include FASHIONGO.ahk
#Include UPS_InputAdd.ahk
#Include SHOWPHONEEMAIL.ahk
#Include ApplyCreditFunction.ahk
#Include CConLAMBS.ahk
#Include GUI_UserNameAndNumberOfDisapproval.ahk
#Include GetInfoFromFashiongo.ahk
#Include GetInfoFromJODIFL_WEB.ahk
#Include GetInfoFromLASHOROOM.ahk
#Include 1stCommonLAMBSProcessing.ahk

;이미지 검색을 위한 전역변수 선언
global pX, pY, jpgLocation 

;주소 등을 넣을 전역변수 선언
global  CompanyName, Attention, Address1, Address2, ZipCode, City, Phone, Email, SubTotal, Invoice_Memo

global TrackingNumber, UserName, Decline1st, Decline2nd, Decline3rd, InvoiceBalance




;Invoice_Memo = invoice memo ;인보이스 메모 내용 담기 위한 변수

WindowName = "" ;활성화 시킬 윈도우 제목 넣는 변수

InvoiceMemoOnLAMBS = "" ;LAMBS의 Invoice Memo 내용 저장하는 변수

CustomerMemoOnLAMBS = "" ;LAMBS의 Customer Memo 내용 저장하는 변수

	F_arr := [] ;패션고 PO넣을 배열
	L_arr := [] ;웹     PO넣을 배열
	W_arr := [] ;LA쇼룸 PO넣을 배열
	
	i = 1 ;패션고 배열을 위한 카운터 변수
	j = 1 ;웹     배열을 위한 카운터 변수
	k = 1 ;LA쇼룸 배열을 위한 카운터 변수
	
	lv_F = 1 ;패션고 PO의 마지막 위치 저장하는 변수
	lv_W = 1 ;웹     PO의 마지막 위치 저장하는 변수
	lv_L = 1 ;LA쇼룸 PO의 마지막 위치 저장하는 변수
	
	Invoice_Memo = "" ; Invoice Memo 내용 저장하는 변수
	FoundPos = 1
	
	Box_arr := [] ;박스가 한 개 이상일 때 넣을 변수 선언
	l = 1 ;박스 갯수를 위한 카운터 변수

	
	OrdersFrom = 0
	


	


;GUI Backgroud
;Gui, Show, w327 h137, JODIFL
Gui, Show, w350 h230, JODIFL

;Input invoice Number
;Gui, Add, Text, x22 y19 w70 h20 , Invoice No.
Gui, Add, Text, x22 y19 Cred , Invoice No.
Gui, Add, Edit, x92 y19 w100 h20 vInvoice_No,  70829 ;67629 ;Puerto Rico Order 72253

;Input Wts of boxes
;Gui, Add, Text, x22 y49 w70 h20 , Wts
Gui, Add, Text, x22 y49 CBlue , Wts
Gui, Add, Edit, x92 y49 w100 h20 vInvoice_Wts, 22

;엔터 버튼
Gui, Add, Button, x225 y19 w100 h80 gClick_btn, Enter

;Input # of boxes
Gui, Add, Text, x22 y79 w70 h20 , No. of Box
Gui, Add, Edit, x92 y79 w100 h20 vH1,
Gui, Add, UpDown, x172 y79 w20 vNo_of_Boxes, 1

;Apply Credit 체크박스
Gui, Add, CheckBox, x22 y115 w80 h30 vApplyCredit +Center, Apply Credit


;Consolidation(Delivery) 체크박스
Gui, Add, CheckBox, x110 y115 w100 h30 vConsolidation +Center, Consolidation`n(Delivery)

;Customer's UPS Account
Gui, Add, Text, x220 y120 w100 h40 +Center, Customer's UPS Account
Gui, Add, Edit, x222 y151 w100 h20 vCustomerUPSAccount, ;X870y4

;2nd Day 체크박스
Gui, Add, CheckBox, x22 y150 w80 h30 v2ndDay +Center, 2nd Day

;3rd Day 체크박스
Gui, Add, CheckBox, x110 y150 w80 h30 v3rdDay +Center, 3rd Day

;Next Day 체크박스
Gui, Add, CheckBox, x22 y185 w80 h30 vNextDay +Center, Next Day

;GUI시작 시 포커스를 Invoice_No 입력칸에 위치
GuiControl, Focus, Invoice_No


return

;GuiClose:
;ExitApp

Click_btn:

;	Gui, Submit

	; To use the values which input on GUI
	Gui Submit, nohide
	
	GUI, Destroy
	
	
	1stCommonLAMBSProcessing(Invoice_No, Invoice_Wts, No_of_Boxes, ApplyCredit, Consolidation, CustomerUPSAccount, NextDay, 2ndDay, 3rdDay)


/*	
	;화면 초기화
	Start()
	Start_Invoice_1()

	;Create Invoice로 돌아가기
	OpenCreateInvoiceTab()

	sleep 1000


	;New & Clear 버튼 클릭
	MouseClick, l, 60, 124, 2, 
	sleep 1000
	
	;Order ID 라디오 버튼 선택하기
	MouseClick, l, 37, 205, 2, 
	sleep 500
	
	;Invoice 입력하기
	MouseClick, l, 209, 205, 2, 
	SendInput, %Invoice_No%
	SendInput, {Enter}
	sleep 1000
	
	
	; [예외처리] Invoice 번호 잘못 입력하여 경고창 나왔을 때 엔터 누른 뒤 다시 시작하기
	IfWinExist, Confirm
	{
		SendInput, {Enter}
		reload
	}
	
	;Security 버튼 있는지 찾기
	;그림이 찾아졌으면 클릭
	;비밀번호 입력칸에 1 입력 후 엔터
	FindSecurityButtonAndClickItThenInputNumber1()


	; Invoice 1 탭에서 시작하기 위해 클릭
	MouseClick, l, 45, 263, 1

	;Customer 클릭해서 정보 갱신하기
	MouseClick, l, 306, 297
	MouseClick, l, 270, 297
	SendInput, {Enter}
	sleep 1000


	;무게 입력하기
	MouseClickDrag, l, 591, 476, 363, 476
	sleep 500
	SendInput, %Invoice_Wts%
	SendInput, {Enter}


	;상자 갯수 입력하기
	MouseClickDrag, l, 591, 500, 363, 500
	sleep 500
	SendInput, %No_of_Boxes%
	SendInput, {Enter}


	;인보이스 메모 처리
	;맨 아랫줄 끝으로 가서 TRACKING # 입력
	Clipboard := 
;	MsgBox, Clipboard`n%Clipboard%
	SendInput, {down 50}
	SendInput, {End}
	Send, {SPACE}TRACKING{#}{SPACE}

	SendInput, ^a^c
	sleep 1000 ;클립보드에 내용 복사 후 꼭 1초를 쉬어줘야 제대로 클립보드에 입력됨
	
	;MsgBox, %clipboard%
	InvoiceMemoOnLAMBS := % Clipboard
;	MsgBox, InvoiceMemoOnLAMBS`n`n`n%InvoiceMemoOnLAMBS%
	;sleep 1000
	
	
	;Customer Memo 읽어오기
	Clipboard := 
;	MsgBox, Clipboard`n%Clipboard%
	MouseClick, l, 565, 626, 2
	
	;sleep 1000
	
	SendInput, ^a^c
	sleep 1000 ;클립보드에 내용 복사 후 꼭 1초를 쉬어줘야 제대로 클립보드에 입력됨
	
	CustomerMemoOnLAMBS := % Clipboard
;	MsgBox, CustomerMemoOnLAMBS`n`n`n%CustomerMemoOnLAMBS%
	
	
	;아이템 갯수 확인 위해 Details Edit열기
	SendInput, !d
	
	
	
	;Details Edit창이 없어질때까지 대기
	CheckTheWindowPresentAndWaitUntillItClose("Details Edit")
	Sleep 500
	
	;저장 후 Status창이 나타나면 없어질때까지 대기
	IfWinExist, Status
		CheckTheWindowPresentAndWaitUntillItClose("Status")


	Sleep 500

	;Security 버튼 있는지 찾기
	;그림이 찾아졌으면 클릭
	;비밀번호 입력칸에 1 입력 후 엔터
	FindSecurityButtonAndClickItThenInputNumber1()
	
	
	;아이템을 추가할 것인지 묻고 Yes 눌렀으면 Sales Orders 버튼 클릭
	MsgBox, 260, Add Items to, Would you like to transfer Items from Sales Order?
	IfMsgBox, Yes
	{
		MouseClick, l, 232, 388
		
		;Transfer from Sales Order창이 없어질때까지 대기
		CheckTheWindowPresentAndWaitUntillItClose("Transfer from Sales Order")
		
		;Security 버튼 있는지 찾기
		;그림이 찾아졌으면 클릭
		;비밀번호 입력칸에 1 입력 후 엔터
		FindSecurityButtonAndClickItThenInputNumber1()
	}
*/

		; 만약 Consolidation 이 체크 됐으면 CC 함수 호출
		if(Consolidation){
			
			; 각 주문별 CC함수 호출
			if(loc_of_MostRecentPo = 1)
				CConLAMBS()
			else if(loc_of_MostRecentPo = lv_F)
				CConFashiongo()
			else if(loc_of_MostRecentPo = lv_W)
				CConJODIFL_WEB()
			else
				CConLASHOROOM()
			
		; 카드 결제가 제대로 됐는지 묻고 Yes 눌렀으면 UPS Label 인쇄 함수 호출
		MsgBox, 4, UPS Label Print out, Does Credit Card Go Passed?
		IfMsgBox, Yes
		{
			UPSLabelPrintOut()
			WrapUp1st(Consolidation)
			Reload

		}
		else IfMsgBox, No
		{
			;Decline 이메일 보내는 함수 호출
			MsgBox, NO~~~
			DeclineProcessing()
			Reload
			
		}	
	}
	

	
	
	
	
	
;	MsgBox, Set the searching status of FASHIONGO

	Invoice_Memo := % InvoiceMemoOnLAMBS
	;Invoice_Memo = , Sales #41668/PO #MTR1B08BD00EA, Sales #42218/PO #MTR1B1FF6DECF TRACKING# SENT EMAIL -EUNICE 6/5/17 LVM&LEM 6/8/17 KRISLEN // LVM 06/12/2017, SENT EMAIL-JAZMINE//   Customer was contacted through phone call + email and no response was received. Order cancelled due to lack of payment. KRISLEN** 6/14/2017. TRACKING #
	
;	MsgBox, % Invoice_Memo
	
	sleep 1000
	
	;인보이스 내용에서 각 주문의 PO만 읽는 함수
	;세 곳의 po가 없는데도 (쇼,전화,이메일 주문일 경우)
	;FoundPos값이 처음에 1인 것 때문에 패션고 실행으로 넘어가는데 이거 고쳐야 됨
	OrdersFrom := 
	FindingPOs(Invoice_Memo, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L)
	
	/*
	;MsgBox, % i
	MsgBox, % FoundPos
	MsgBox, % F_arr[i]
	MsgBox, % W_arr[j]
	MsgBox, % L_arr[k]
	
	MsgBox, lv_F`n%lv_F%
	MsgBox, lv_W`n%lv_W%
	MsgBox, lv_L`n%lv_L%
	*/
	
	;배열값은 함수로 직접 넘겨주지 못하더라. 그래서,
	F_PO = % F_arr[i] ;패션고의 PO중 가장 최근의 PO값을 F_PO에 넣었다.
	W_PO = % W_arr[j] ;웹의     PO중 가장 최근의 PO값을 W_PO에 넣었다
	L_PO = % L_arr[k] ;LA쇼룸의 PO중 가장 최근의 PO값을 L_PO에 넣었다
	
	
;	MsgBox, F_PO`n%F_PO%
;	MsgBox, W_PO`n%W_PO%
;	MsgBox, L_PO`n%L_PO%
	
	
	
	;숫자가 작을 수록 먼저 찾은, 즉 오래된 PO이고 숫자가 가장 높은 것이 가장 나중에 찾은, 즉 가장 최근 PO니까
	;각 주문처의 PO가 찾아진 FoundPos값 중 가장 나중의 값이 저장된 (패션고:lv_F   웹:lv_W   LA쇼룸:lv_L)
	;위의 세 변수 중 가장 숫자가 큰 것이 모든 것 통틀어 가장 최근의 PO이다.
	loc_of_MostRecentPo := get_max_among_3(lv_F, lv_W, lv_L)
;	MsgBox, Most Recent Po# is`n%loc_of_MostRecentPo%
	

	;loc_of_MostRecentPo 변수에는
	;세 곳의 각 마지막 FoundPos값(PO가 가장 나중에 찾아진 위치)
	;가운데 가장 큰 수(즉, 가장 최근의 PO값이 찾아진 위치)가 들어있는데
	;loc_of_MostRecentPo 변수와 맞는 값을 찾아서(그럼 이게 가장 최근의 PO니까)
	;그 주문 웹 페이지로 이동하는 함수를 또 호출 한다
	
	;FoundPos값을 검색을 위해 처음에 1로 설정하는데
	;이것 때문에 세 곳의 PO를 못 찾아도(전화,이메일,쇼 주문)
	;lv_F lv_W lv_L 값도 1이 되고 결국
	;마지막에 loc_of_MostRecentPo값도 1이 되어
	;Active_and_Find_PO_in_Fashiongo함수가 실행된다.
	;이것을 방지하기 loc_of_MostRecentPo값이 1이면 
	;웹주문 페이지로 넘어가는 것을 방지하기 위핸 예외처리 처음에 넣어줬다

	if(loc_of_MostRecentPo = 1){
;		MsgBox, 쇼,전화,이메일 주문입니다. ;FoundPos: `n%FoundPos%
		
		;웹에서 온 주문이 아니기 때문에 LAMBS에서 정보 읽고 저장하는 함수 호출
		SHOWPHONEEMAIL()
;		MsgBox, Out of SHOWPHONEEMAIL
		
		;return
	}
	else if(loc_of_MostRecentPo = lv_F)
		GetInfoFromFashiongo(F_PO)
	else if(loc_of_MostRecentPo = lv_W)
		GetInfoFromJODIFL_WEB(W_PO)
	else
		GetInfoFromLASHOROOM(L_PO)
	
	
	; LAMBS의 인보이스메모, 커스토머메모, 웹의 커스토머메모 띄우기
	MsgBox, Invoice Memo On LAMBS`n%InvoiceMemoOnLAMBS%`n`n`nCustomer Memo On LAMBS`n%CustomerMemoOnLAMBS%`n`n`nCustomer Notes on Web`n	
	
	

	
/*	
	; 1박스 이상이면 배열에 넣기
	if(No_of_Boxes >= 2){
		Box_arr[l] = 
		MouseClickAndPaste(338, 534, Weight)
			
	}
	
	if(strLen(SubPat2) >= 10){
		F_arr[i] := SubPat2
		i += 1
		lv_F := % FoundPos
		MsgBox, lv_F`n%lv_F%
	}
*/

	
	; UPS에 주소 입력하기 위해 함수 호출
	; RoundedShippingFee에는 올림 된 배송비 저장
	RoundedShippingFee := UPS_InputAdd(NextDay, 3rdDay, 2ndDay, CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts)
;	MsgBox, % RoundedShippingFee


	;화면 초기화
	Start()
	Start_Invoice_1()
	
	; EST/Actual FC에 배송비 입력하기	
	DragAndPast(810, 336, 730, 336, RoundedShippingFee)
	SendInput, {Tab}
	
	; Invoice Balance 값 저장하기
	InvoiceBalance := ClickAndCtrlAll(867, 591)
	SendInput, {F8}

	;security 버튼 찾아서 풀어주기
	FindSecurityButtonAndClickItThenInputNumber1()
	
	; 만약 ApplyCredit 가 체크 됐으면 ApplyCredit함수 호출
	if(ApplyCredit){
		
		ApplyCreditFunction(Invoice_No)
		Sleep 500
	}

/*
	; 쇼, 전화, 이메일 주문이라면 LAMBS에서 CC 창 열기
	if(loc_of_MostRecentPo = 1){
		CConLAMBS()
;		MsgBox, cc WINDOW oUT
	}
*/
	; 각 주문별 CC함수 호출
	if(loc_of_MostRecentPo = 1)
		CConLAMBS()
	else if(loc_of_MostRecentPo = lv_F)
		CConFashiongo()
	else if(loc_of_MostRecentPo = lv_W)
		CConJODIFL_WEB()
	else
		CConLASHOROOM()






	; 카드 결제가 제대로 됐는지 묻고 Yes 눌렀으면 UPS Label 인쇄 함수 호출
	MsgBox, 4, UPS Label Print out, Does Credit Card Go Passed?
	IfMsgBox, Yes
	{
		UPSLabelPrintOut()
		
		;LAMBS의 Invoice Memo에 Tracking Number 적어넣는 함수호출
		WrapUp1st(Consolidation)
		
		Reload
/*		
		;LAMBS에 Tracking Number 적어넣기
		Start()
		FindSecurityButtonAndClickItThenInputNumber1()
		MouseClick, l, 602, 498
		SendInput, {Tab}
		SendInput, % TrackingNumber
		SendInput, {F8}
		FindSecurityButtonAndClickItThenInputNumber1()
*/		
	}
	else IfMsgBox, No
	{
		;Decline 이메일 보내는 함수 호출
		MsgBox, NO~~~
		DeclineProcessing()
		Reload
		
	}


	/*
	; consolidation이면 이리로 점프해서 온다
	Saving:
		WrapUp1st()
		Reload
	*/
	
	
		
/*	
	Control_InputText("WindowsForms10.EDIT.app.0.378734a71", RoundedShippingFee, %windowtitle%)
	Control_SnedButton("F8", %windowtitle%)
*/	
	;ControlSend, Edit9, {Enter}, %windowtitle%
	
	;WindowsForms10.EDIT.app.0.378734a71
	

	;UPS_InputAdd(CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts)

	
	
	
	; 프로그램 처음 시작 시 Gui 이요하여패션고 검색 세팅(PO Number로 바꾸고 Select Period로 바꿔야 되는 것 표시하기) 처음 시작 한 번만
	
	; 가장 최근의 po찾은 다음에 LAMBS에서 아이템 확인 후(WinExist써서 아이템 에디트 창이 사라지면 시큐리티 버튼 찾고 찾으면 1눌러주고 없으면 넘어가서)
	; 그 다음에 가장 최근의 po가 있는 웹 주문 페이지로 이동

	; LA쇼룸 창 열 때마다 로그아웃 됐는지 확인하는 함수 만들어서 그때마다 확인(심지어 업무 중간에 로그아웃 되기도 한다)
	; 아이템 갯수 수정 후 Status 창이 나오면 기다렸다가 다음으로 넘어가도록 하자 sleep 쓰는 대신에
	; LAMBS에서 Status 창이 나오면 기다리자. 그런데 프린터가 꼬져서 Invoice(M) 눌렀는데 한참 있다가 Status창 뜨기도 했다

	; UPS 경고창이 뜨면 그냥 기다리게 하자. 예외처리가 너무 많아지거나 중요 메세지 그냥 넘길 수 있다.
	; UPS 경고창이 뜨면 CONTINUE? 메세지 박스 띄워서 끝날때까지 기다리지
	; UPS LA 지역이면 배송 아닌 배달일 가능성이 대부분이니 확인창 띄우자(전화로 고객에게 확답받도록);
	; LASHOROOM 에러 너무 많이 나는데 어떻게 일일히 에러 처리하지?
	; LASHOROOM Company 이름 바로 밑에 [BLOCKED FROM WEB SITE] 경고창 있는경우도 있었음
	; 인보이스 번호 입력 후 회사 명 다시 선택해서 리프래쉬 해야됨
	; 웹으로 넘어가기 전까지는 마우스/키보드 입력 막는 것 쓸까? 하긴 그러면 아이템 수정을 할 수 없게 된다

	; 백오더가 없는데 Sales Orders 클릭하니 Confirm (No Data) 안내 창이 뜨는데 이거 처리해야됨 

	; 2ND DAY, 3RD DAY UPS처리 해줘야 됨
	; consolidation 처리 해줘야 됨
	
	; decline 됐을 때 이메일 보내는 것 완전히 끝내지 않았음
	; 주문이 추가된 경우에는 나중의 PO가 더 오래된 것일 수 있다. 백오더가 나중에 더해질 수도 있어서 72167
	
	; DeclineProcessing 함수 호출 시작할 때 언어 설정을 영어로 바꾸는 것 해야지 오류 안 나지 않을까
	
	
	
	
	


Exitapp

F7::
Reload
	
Esc::
 Exitapp
 Reload