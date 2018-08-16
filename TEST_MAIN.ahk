#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include UPS_InputAdd.ahk
#Include GetInfoFromLAMBS.ahk
#Include ApplyCreditFunction.ahk
#Include CConLAMBS.ahk
#Include CConFashiongo.ahk
#Include CConLASHOROOM.ahk
#Include GUI_UserNameAndNumberOfDisapproval.ahk
#Include GetInfoFromLAMBS.ahk
#Include GetInfoFromFashiongo.ahk
#Include GetInfoFromJODIFL_WEB.ahk
#Include GetInfoFromLASHOROOM.ahk
#Include 1stCommonLAMBSProcessing.ahk
#Include GetActiveBrowserURL.ahk
#Include OnlinePayment.ahk
;#Include UrlDownloadToVar.ahk


;이미지 검색을 위한 전역변수 선언
global pX, pY, jpgLocation 

;주소 등을 넣을 전역변수 선언
global  CompanyName, Attention, Address1, Address2, ZipCode, City, Phone, Email, SubTotal, Invoice_Memo, State, Country, BillingAdd1, BillingZip, RoundedShippingFee, InvoiceBalance

global CCNumbers, CVV, Month, Year, ExpDate, iCountForOnlinePayment

global CCNumbers2, ExpDate2, CVV2, Month2, Year2
global CCNumbers3, ExpDate3, CVV3, Month3, Year3
global CCNumbers4, ExpDate4, CVV4, Month4, Year4

global TrackingNumber, UserName, Decline1st, Decline2nd, Decline3rd, InvoiceBalance, Wts_of_Boxes

global InvoiceMemoOnLAMBS, CustomerMemoOnLAMBS, CustomerNoteOnWeb, StaffOnlyNote, Invoice_No, FGServer, wb, Paymentwb ;, CCinfoOnFASHIONGO



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
	
	
	
	
	
loc_of_MostRecentPo = 1	
	
	
	
	


	;Invoice_Memo = , Sales #43321/PO #MTR171EC8, Sales #43320/PO #PHONE ORDER, Sales #43320/PO #YULIAM 7/7/2017, Sales #43320/PO #TIFFANY, Sales #44310/PO #MTR171F67, Sales #45015/PO #MTR171FC7, Sales #48103/PO #7/21/17 TRACKING#
	Invoice_Memo = , Sales #42782/PO #MTR1B3844B399 TRACKING#
	;Invoice_Memo :=
	
	;Invoice_Memo = , Sales #45954/PO #MTR1BA4F67AA1, Sales #47425/PO #MTR1BD9577950 TRACKING# 1ZW947530364804034
	
	
	;Invoice_Memo = , Sales #37523/PO #MTR171A7E, Sales #40284/PO #OP041517043, Sales #40284/PO #OP072125517, Sales #45954/PO #MTR1BA4F67AA1, Sales #47425/PO #MTR1BD9577950 TRACKING# 1ZW947530364804034



	Email = han1002@daum.net
	
	Invoice_No = 70829 ;70829 
	Invoice_Wts = 11, 22, 33, 44
	No_of_Boxes = 2
	ApplyCredit = 1
	Consolidation = 1
	CustomerUPSAccount = X870y4
	2ndMonth = 1
	3rdMonth = 1
	NextMonth = 1
	
	
	CCNumbers = 123456789
	
	
	
	CompanyName := "WAREHOUSE " ;"CHANTILLY BOUTIQUE"
	
	Attention = KURT SCHOLLA ;111MITZY BURROUGHS
	Address1 = 4056 BROADWAY SUITE # 1 ;3834 CENTRAL AVE SUITE A
	Address2 = SUITE A
	ZipCode = 64111 ;91020 ;71913
	City = KANSAS CITY ;HOT SPRINGS NATIONAL PARK
	State = MO
	Phone = 12139998429 ;501-627-8613
	SubTotal = 1,045.50	 ;DECLINE #1 EMAIL SENT 07/10/2017 asdf 
	RoundedShippingFee = 15
	InvoiceBalance = 1,000.50



	BillingAdd1 = 1945 JODIFL
	BillingZip = 91020
	
	CCNumbers = 123456789	
	CCNumbers = 4000000000000000	
	CVV = 123
	Month = 10
	Year = 2020



	CustomerUPSAccount = ;X870y4
		
	Invoice_No = 70829
	No_of_Boxes = 2
	Invoice_Wts = 22
	NextMonth = 1	

	

;OnlinePayment()


/*
	;UPS_InputAdd 실행하는 부분

	State = P
	MsgBox, % State
	

	;Country = United States
	Country = Canada

	MsgBox, % Country
	

	; 해외 주문 찾아내기
	IfEqual, State, PR
	{
		MsgBox, It's Order fromDDD Puerto Rico. Now converts to manual mode.
		Reload
	}
	
	IfNotEqual, Country, United States
	{
		MsgBox, It's Order from out of country. Now converts to manual mode.
		Reload
	}
	
	if(Country != United States){

	}
	
	else if(%Country% = Puerto Rico){
		MsgBox, It's Order from Puerto Rico. Now converts to manual mode.
		Reload
	}

	

	RoundedShippingFee := UPS_InputAdd(NextMonth, 3rdMonth, 2ndMonth, CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts, NoDeclare)

	MsgBox, Rounded Shipping Fee`n`n%RoundedShippingFee%
*/





/*
;ApplyCreditFunction 실행하는 부분


	Invoice_No = 73933
	ApplyCreditFunction(Invoice_No)
*/

/*
;CC 실행하는 부분

	Address1 = 3834 CENTRAL AVE SUITE A
	ZipCode = 71913
	CC()
	MsgBox, Out!
*/


/*
;ups 라벨 인쇄하는 부분

	; 카드 결제가 제대로 됐는지 묻고 Yes 눌렀으면 UPS Label 인쇄 함수 호출
	MsgBox, 4, UPS Label Print out, Does Credit Card Go Passed?
	IfMsgBox, Yes
	{
		UPSLabelPrintOut()

	}
	else IfMsgBox, No
	{
		;Decline 이메일 보내는 함수 호출
		DeclineProcessing()
		MsgBox, NO~~~
	}
	
	
	MsgBox, % TrackingNumber
	;GUI_TrackingNumber()
*/	




/*
	windowtitle = Untitled - Message (HTML)
	
	WinWait, %windowtitle%
	
	Control_InputText("RichEdit20WPT1", Email, windowtitle)
	
	
	FormatTime, CurrentDateTime, ,MM/dd/yyyy
	title = CREDIT CARD DECLINE NOTIFICATION %CurrentDateTime% JODIFL
	
	Control_InputText("RichEdit20WPT4", title, windowtitle)
	
	ControlSend, , {!}, windowtitle
	ControlSend, , {n}, windowtitle
	ControlSend, , {a}, 1
	
	ControlSend, , {s}, windowtitle	
*/
/*	
	ControlSend, NetUIHWND1, {!}, %windowtitle%
	ControlSend, NetUIHWND1, {n}, %windowtitle%
	ControlSend, NetUIHWND1, {a}{s}, %windowtitle%
*/



	

;	emailSearch()
/*	
	WinActivate, LAMBS
	MouseMove, 938, 251, 70
	Sleep 100
	Click
*/



/*
	; Email 체크 버튼 누르기
	WinActivate, LAMBS
	MouseMove, 935, 251
	MouseGetPos, , , , control	
	ControlSend, %control%, {Space}, LAMBS	
*/

;	WinClose, vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR2C9F3098E - Google Chrome

;	ApplyCreditFunction(72383)

;	UPSLabelPrintOut()



/*	if(WinWait, UPS WorldShip, , 1){
		WinActivate
		ControlSend, Button1, {Enter}, UPS WorldShip
	}
*/
	

;		Control_SnedButton(F10, windowtitle) ;이걸로 한 번 해보자
		;SendInput, {F10}


;	WrapUp1st1


	



;GetInfoFromLASHOROOM("OP073026207")
;MsgBox, PAUSE LASHOWROOM

GetInfoFromFashiongo("MTR2CE7A4FB4", 2)
MsgBox, fashiongo pause
/*
GetInfoFromFashiongo("MTR2CE7A4FB4", 2)
UPS_InputAdd(NextMonth, 3rdMonth, 2ndMonth, CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts, NoDeclare)
MsgBox, pause lashoroom
*/



;1


;		URL = http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR2B4FC1B2C
		; 백 오더에는 po번호에 -BO가 붙는다 BO BO1 BO2 이런 식으로
		
;		URLDownloadToFile, %URL%, 1.htm

;	SHOWPHONEEMAIL()

/*
	WinActivate, LAMBS
	jpgLocation = %A_ScripDir%PICTURES\Qty.png

	PicSearch(jpgLocation)
*/


;	CoordMode, mouse, Relative
	;CoordMode, mouse, Screen
	
	
;	WinActivate, UPS WorldShip - Remote Workstation
	


;	MouseMove, 1243, 114
/*
	;F10 Click
	MouseMove, 575, 627, 50
	click			
	Send, !{tab}


	;F2, F3 Click
;	Send, !{tab}	
	MouseMove, 37, 84, 50
	click
	Send, !{tab}	
	Sleep 5000
	Send, !{tab}	
	
	MouseMove, 127, 250, 50
	MouseClick, l, 127, 250, , , d
	Send, {enter}
	Click
	Sleep 100
	
	Click
	
	MouseMove, 127, 264, 50
	Click
	Sleep 100	
	Click			
*/
			
			
			;Send, {F2}
			
			
/*
	; 마우스를 새 위치로 이동합니다:
	MouseMove, 200, 100

	; 마우스를 천천히 (속도 50 vs. 2) 현재 위치로부터
	; 20 픽셀만큼 오른쪽으로 그리고 30 픽셀 만큼 아래쪽으로 이동시킵니다:
	MouseMove, 20, 30, 100, R
*/


/*
	MouseGetPos, xpos, ypos 
	Msgbox, The cursor is at X%xpos% Y%ypos%. 

	; 이 예제에서 마우스를 이동시켜서 현재 마우스 아래에 있는
	; 창의 제목을 볼 수 있습니다:
	#Persistent
	SetTimer, WatchCursor, 100
	return

	WatchCursor:
	MouseGetPos, , , id, control
	WinGetTitle, title, ahk_id %id%
	WinGetClass, class, ahk_id %id%
	ToolTip, ahk_id %id%`nahk_class %class%`n%title%`nControl: %control%
	return
*/
	
;	1stCommonLAMBSProcessing(Invoice_No, Invoice_Wts, No_of_Boxes, ApplyCredit, Consolidation, CustomerUPSAccount, NextMonth, 2ndMonth, 3rdMonth)



/*
	PO_L = OP032115237
	GetInfoFromLASHOROOM(PO_L)
*/



/*
	;StringSplit 예제. %A_Space%의 쓰임
	TestString = This is a test.
	
	; 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 word_array에 저장
	StringSplit, word_array, TestString, `,|%A_Space%, .  ; 점은 제외합니다.

	MsgBox, The 4th word is %word_array4%.

	Colors = red,green,blue
	StringSplit, ColorArray, Colors, `,
	Loop, %ColorArray0%
	{
		this_color := ColorArray%a_index%
		MsgBox, Color number %a_index% is %this_color%.
		NumofVal := % A_Index
	}
	
	MsgBox, The number of value of ColorArray is `n`n%NumofVal%
*/





	;MsgBox, % No_of_Boxes

/*
	; Invoice_Wts값에 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 Wts_of_Boxes에 저장
	StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.


	; Wts_of_Boxes를 루프 돌려서 들어있는 값 개수만큼 No_of_Boxes에 저장
	Loop, %Wts_of_Boxes0%{
		this_Wts := Wts_of_Boxes%a_index%
	;	MsgBox, The Weight Box number %a_index% is %this_Wts%
	;	No_of_Boxes := % A_Index

		;if(%a_index% < 4){
		if(%a_index% < 4){
	;		MsgBox, plus %a_index%
		}
	}

;	MsgBox, The number of value of ColorArray is `n`n%No_of_Boxes%
*/
/*
windowtitle = UPS WorldShip - Remote Workstation
/*	; 상자 개수가 1개 이상이면
	if (1 <= No_of_Boxes){
		Loop, %Wts_of_Boxes0%{
			this_Wts := Wts_of_Boxes%a_index%
			;MsgBox, The Weight Box number %a_index% is %this_Wts%			
			Control_InputText("Edit32", this_Wts, windowtitle)
;			Sleep 1000
			ControlClick, Button29, %windowtitle%
			Sleep 2000
		}
	ControlClick, Button3, %windowtitle%
	}
	else
		Control_InputText("Edit32", this_Wts, Invoice_Wts)
*/	
/*
	; 무게 입력
	if (1 == No_of_Boxes){
		Control_InputText("Edit32", this_Wts, windowtitle)
		MsgBox, 111
	}
	else{
		Loop, %Wts_of_Boxes0%{
			this_Wts := Wts_of_Boxes%a_index%
			;MsgBox, The Weight Box number %a_index% is %this_Wts%			
			Control_InputText("Edit32", this_Wts, windowtitle)
;			Sleep 1000
;			ControlClick, Button29, %windowtitle%
			ControlClick, x325 y642, %windowtitle%
			Sleep 2000
		}
;	ControlClick, Button3, %windowtitle%
	ControlClick, x371 y619, %windowtitle%
	}
*/			

	


/*	
	
	; 무게 입력
	; 상자 개수가 1개 이상이면
	if (1 == No_of_Boxes)
		Control_InputText("Edit32", Invoice_Wts, windowtitle)
	else{
		; Invoice_Wts값에 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 Wts_of_Boxes에 저장
		StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.

		; Wts_of_Boxes를 루프 돌려서 들어있는 값 개수만큼 No_of_Boxes에 저장
		Loop, %Wts_of_Boxes0%
		{
			this_Wts := Wts_of_Boxes%a_index%
		;	MsgBox, The Weight Box number %a_index% is %this_Wts%
		;	No_of_Boxes := % A_Index
		}
	}
	
*/	

/*
	; 패션고 정보가 LAMBS와 맞는지 묻고 Yes 눌렀으면 LAMBS에서 정보 얻는 SHOWPHONEEMAIL() 함수 호출
	MsgBox, 4096, UPS Label Print out, 고객 메모 확인`n`nIs Info on FASHIONGO same as LAMBS?
	IfMsgBox, Yes
	{
		SHOWPHONEEMAIL()
		

	}
	else IfMsgBox, No
	{
		Reload
		
	}
*/
/*
	GetInfoFromFashiongo(PO_F)

	MsgBox, 4100, UPS Label Print out, 고객 메모 확인`n`nIs Info on FASHIONGO same as LAMBS?
*/

/*
SubStr(String, StartingPos [, Length]) [v1.0.46+]: 부분문자열을 String으로부터 복사합니다. StartingPos에서 시작해서 오른쪽으로 진행해서 최대 Length개의 문자를 포함합니다 (Length를 생략하면, 기본 값은 "모든 문자"입니다). StartingPos에 1을 지정하면 첫 번째 문자에서, 2을 지정하면 2번째 문자부터 시작합니다. 등등. ( StartingPos가 String의 길이를 넘어서면, 빈 문자열이 반환됩니다). StartingPos가 1보다 작으면, 문자열 끝으로부터의 오프셋으로 간주됩니다. 예를 들어, 0은 가장 마지막 문자를 추출하고, -1이면 그 문자열의 마지막 문자로부터 1만큼 왼쪽으로 떨어져 있다고 간주됩니다. (그러나 StartingPos가 문자열의 왼쪽 끝을 넘어서 시도하면, 첫 문자부터 추출을 시작합니다). Length는 열람할 최대 문자 개수입니다 (문자열의 나머지 부분이 너무 짧으면 최대 개수보다 적게 열람됩니다). 음의 길이(Length)를 지정하면 반환된 문자열의 끝으로부터 문자를 그 개수 만큼 생략합니다 (모든 또는 너무 많은 문자를 생략하면 빈 문자열이 반환됩니다). 관련 항목: RegExMatch(), StringMid, StringLeft/Right, StringTrimLeft/Right.
*/

/*
4802
1376
0184
9211
*/

/*
	StartingPos = 1

	loop, 5{
		PartiallyCCNum%A_Index% := SubStr(CCNumbers, StartingPos, 4)
		;MsgBox, % PartiallyCCNum%A_Index%
		StartingPos := StartingPos + 4
	}

	MsgBox, %PartiallyCCNum1%    %PartiallyCCNum2%    %PartiallyCCNum3%    %PartiallyCCNum4%    %PartiallyCCNum5%
	;MsgBox, % PartiallyCCNum2
*/

;CConLAMBS(loc_of_MostRecentPo)

;CallCConLAMBSAndAskWhetherCCwentOrNot()



		
/*		
;		MsgBox, % OrderIdVal

		MouseMove, 904, 305
		MouseGetPos, , , , control
		ControlGetText, OrderIdVal, %control%, LAMBS
		MsgBox, % OrderIdVal
		
		MouseMove, 904, 285
		MouseGetPos, , , , control
		ControlGetText, OrderIdVal, %control%, LAMBS		
		MsgBox, % OrderIdVal
*/



/*
		WinActivate, LAMBS
		MouseMove, 904, 325
		MouseGetPos, , , , control
		ControlGetText, OrderIdVal, %control%, LAMBS
*/		
		
		


;		WinActivate, LAMBS
/*		
		; 주황색 [Credit] 체크박스 체크하기
		MouseMove 665, 892
		Sleep 200		
		MouseMove 665, 912
		Sleep 200				
*/		
/*		
		; 파랑색 [Credit] 체크박스 체크하기
		Clipboard := 
		
		; 맨 위의 Order Id 값 얻기
		MouseMove, 831, 278
		SendInput, #q
		Sleep 100
		MouseMove, 882, 292
		Sleep 100
		SendInput, #q
		Sleep 100
		
		; Capture2Text 창 닫기
		IfWinExist, Capture2Text - OCR Text
			WinClose


		OrderId := % Clipboard
		
		WinActivate, LAMBS
		
		; 파란색 창의 OrderId 값과 Invoice_No가 같으면 체크박스 클릭
		if(OrderId == Invoice_No)
			MouseClick, l, 665, 284
		else{
			MsgBox, 첫째 줄의 Order Id 와 Invoice_No가 맞지 않습니다
			Reload
		}
			
		
		Sleep 200
		
		MsgBox, % OrderId
*/

;ApplyCreditFunction(Invoice_No)



;	DeleteAdd2InAdd1(Address1)

/*
	; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
	DeleteAdd2InAdd1(ByRef Address1){
		return
	}	
*/

/*
	MsgBox, % Invoice_Memo
	
	FindingPOsOfShowPhoneEmail(Invoice_Memo)
*/

	
	
;	SendingEmailCustomerInvoice(Invoice_No)

;	DeclineProcessing()


;	WinClose, view-source:https://admin.lashowroom.com/order_edit_v1.php?order_id=25498&list_option=new - Google Chrome


;  크롬 새창에서 열기

/*	
	Clipboard :=

	url = http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR1BE9C4B535

	run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) url

	WinActivate, vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR1BE9C4B535 - Google Chrome

	Sleep 2000
	SendInput, ^u

	Sleep 1500

	SendInput, ^a^c

	Sleep 1000

	HTMLSource := % Clipboard

	WinClose, view-source:vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR1BE9C4B535 - Google Chrome

	MsgBox, % HTMLSource
*/	

/*
	PO_F = MTR1BE9C4B535
	GetInfoFromFashiongo(PO_F)
;	MsgBox, CompanyName : `n`n%CompanyName%
;	MsgBox, Attention : `n`n%Attention%

*/



/*

; 주문들의 마지막 기본 페이지 열기


	Invoice_Memo = , Sales #37523/PO #MTR171A7E, Sales #40284/PO #OP041517043, Sales #40284/PO #OP072125517, Sales #45954/PO #MTR1BA4F67AA1, Sales #47425/PO #MTR1BD9577950 TRACKING# 1ZW947530364804034
	;Invoice_Memo = , Sales #45954/PO #MTR1BA4F67AA1, Sales #47425/PO #MTR1BD9577950 TRACKING# 1ZW947530364804034
	;Invoice_Memo :=
	

	MsgBox, % Invoice_Memo
	OrdersFrom := 
	FindingPOs(Invoice_Memo, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L)
	
	
	;배열값은 함수로 직접 넘겨주지 못하더라. 그래서,
	F_PO = % F_arr[i] ;패션고의 PO중 가장 최근의 PO값을 F_PO에 넣었다.
	W_PO = % W_arr[j] ;웹의     PO중 가장 최근의 PO값을 W_PO에 넣었다
	L_PO = % L_arr[k] ;LA쇼룸의 PO중 가장 최근의 PO값을 L_PO에 넣었다


	;숫자가 작을 수록 먼저 찾은, 즉 오래된 PO이고 숫자가 가장 높은 것이 가장 나중에 찾은, 즉 가장 최근 PO니까
	;각 주문처의 PO가 찾아진 FoundPos값 중 가장 나중의 값이 저장된 (패션고:lv_F   웹:lv_W   LA쇼룸:lv_L)
	;위의 세 변수 중 가장 숫자가 큰 것이 모든 것 통틀어 가장 최근의 PO이다.
	loc_of_MostRecentPo := get_max_among_3(lv_F, lv_W, lv_L)
	
	MsgBox, i:  %i%
	MsgBox, % F_arr[2]
	
	
	
	; 주문들의 마지막 기본 페이지 열기
	OpenAllBaseWebPageOfOrders(i, j, k, F_arr, W_arr, L_arr)
*/


/*
	; 변수 선언과 사용
	AA := []
	AA[0] := 1
	MsgBox, % AA[0]
*/



/*		
	; LA쇼룸 페이지에서 정보 읽어오기
	FindInfoInLASHOWROOM(HTMLSource){
		
;		FoundPos = 1

		; CompanyName 찾기
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipCompanyName.>(.*)</span>)", SubPat)
		
		; CompanyName 찾았으면 전역변수인 CompanyName에 값 넣기
		if(FoundPos){
			CompanyName := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 CompanyName`n%CompanyName%
			;MsgBox, % FoundPos
		}
		
;		SubPat :=
	
		; Attention 찾기. CompanyName 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAttention.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; Attention 찾았으면 전역변수인 Attention 에 값 넣기
		if(FoundPos){
			Attention := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Attention`n%Attention%
			;MsgBox, % FoundPos
		}
		
		; Address1 찾기. Attention 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; Address1 찾았으면 전역변수인 Address1 에 값 넣기
		if(FoundPos){
			Address1 := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Address1`n%Address1%
			;MsgBox, % FoundPos
		}		
		
		; City 찾기. Address1 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress2.>(.*)[\,])", SubPat, FoundPos + strLen(SubPat))
		
		; City 찾았으면 전역변수인 City 에 값 넣기
		if(FoundPos){
			City := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 City`n%City%
			;MsgBox, % FoundPos
		}
		

		; State 찾기. 처음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress2.>.*(\w\w)\s\d.*</span>)", SubPat) ;, FoundPos + strLen(SubPat))
		
		; State 찾았으면 전역변수인 State 에 값 넣기
		if(FoundPos){
			State := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 State`n%State%
			;MsgBox, % FoundPos
		}

		
		; ZipCode 찾기. 처음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipAddress2.>.*(\d.*)</span>)", SubPat) ;, FoundPos + strLen(SubPat))
		
		; ZipCode 찾았으면 전역변수인 ZipCode 에 값 넣기
		if(FoundPos){
			ZipCode := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 ZipCode`n%ZipCode%
			;MsgBox, % FoundPos
		}
		

		; Country 찾기. ZipCode 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblShipCountry.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; City 찾았으면 전역변수인 City 에 값 넣기
		if(FoundPos){
			Country := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Country`n%Country%
			;MsgBox, % FoundPos
		}		
		
		
		; Phone 찾기. Country 다음부터 찾기 시작
		FoundPos := RegExMatch(HTMLSource, "mU)(lblPhoneShipping.>(.*)</span>)", SubPat, FoundPos + strLen(SubPat))
		
		; Phone 찾았으면 전역변수인 Phone 에 값 넣기
		if(FoundPos){
			Phone := SubPat2
			
			;MsgBox, 함수 안에서 찾은 것 Phone`n%Phone%
			;MsgBox, % FoundPos
		}
		


;		WinActivate, LAMBS
		
		Start()
		Start_Invoice_1()		
		
		MouseMove, 857, 313
		MouseGetPos, , , , control
		ControlGetText, SubTotal, %control%, LAMBS
		
		
		; Invoice 2 탭 클릭하기
		MouseClick, l, 101, 266, LAMBS
		Sleep 100
		
		
		MouseMove, 260, 738
		MouseGetPos, , , , control
		ControlGetText, Email, %control%, LAMBS		
		
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 Address2에 저장하기
		FindAdd2InAdd1(Address1, Address2)
		
		; ADD1에 SUITE등 ADD2 주소 있으면 찾아서 삭제하기
		DeleteAdd2InAdd1(Address1)
		
		
		; Invoice 1 탭 클릭하기
		MouseClick,l , 41, 265		
		
		
		return
	}
	
*/


;DeclineProcessing()


/*
;UPSLabelPrintOut()
PO_F = MTR1BEDA5B361
GetInfoFromFashiongo(PO_F)
MsgBox, pause
*/



/*
;Email 버튼이 체크 안 되었는 지 찾아보고 체크 안 되어 있으면 체크하기
;이메일 보낼 때 사용
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked.png
CheckEmailButtonOrReleaseIt(jpgLocation)
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked_Activated.png
CheckEmailButtonOrReleaseIt(jpgLocation)
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Unchecked_Activated2.png
CheckEmailButtonOrReleaseIt(jpgLocation)
*/

/*
;Email 버튼이 체크 되었는 지 찾아보고 체크 되어 있으면 체크를 풀기
;일반 인쇄 시 혹은 이메일 보낸 후 체크버튼 풀 때 사용
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked.png
CheckEmailButtonOrReleaseIt(jpgLocation)
jpgLocation = %A_ScripDir%PICTURES\EmailButton_Checked_Activated.png
CheckEmailButtonOrReleaseIt(jpgLocation)
*/


;CConLAMBS(loc_of_MostRecentPo)

/*
Consolidation = 1
MsgBox, % Consolidation

	if(Consolidation)
		ConsolidationProcessing(Invoice_Memo, Invoice_No, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L)
		
	MsgBox, pause11
*/



/*
loc_of_MostRecentPo = 1
Invoice_Memo = , Sales #42782/PO #MTR1B3844B399 TRACKING#
WinActivate, LAMBS
WrapUp1st(Consolidation, loc_of_MostRecentPo, Invoice_Memo, Invoice_No)


	; 인보이스에서 패션고, JODIFL, LA쇼룸을 못 찾았으면 고객에게 이메일 발송
	; 만약 패션고, JODIFL, LA쇼룸의 인보이스를 찾았으면 인보이스 중에 쇼,전화,이메일 오더가 있는지 확인 후 있으면 고객에게 이메일 발송
	if(loc_of_MostRecentPo = 1)
		SendingEmailCustomerInvoice(Invoice_No)
	else
		FindingPOsOfShowPhoneEmail(Invoice_Memo, Invoice_No)
*/


/*
wwb := ComObjCreate("InternetExplorer.Application")  ;// Create an IE object
wwb.Visible := true                                  ;// Make the IE object visible
wwb.Navigate("http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR47CEFB0C")                   ;// Navigate to a webpage
;wwb.Navigate("https://admin.lashowroom.com/order_edit_v1.php?order_id=26513&list_option=new")                   ;// Navigate to a webpage
*/



/*
url := "http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR47CEFB0C"

; Example: Download text to a variable:
WB := ComObjCreate("WinHttp.WinWB.5.1")
WB.Open("GET", url)
WB.Send()
this_text := WB.ResponseText
html := ComObjCreate("htmlfile")
html.write(this_text)

; Loop through all links and add them to link_list variable with a new line
Loop % html.links.length
  link_list .= html.links[A_Index - 1].href . "`n"

; Certain links for relative and have text like 'about:/services/' replace the about with url
StringReplace, link_list, link_list, about:, %url%, A
msgbox % link_list


MsgBox, pause00
*/






/*

; LASHOWROOM 로그인

Loginname = jodifl
Password = j123456789
URL = https://admin.lashowroom.com/order_edit_v1.php?order_id=26513&list_option=new

WB := ComObjCreate("InternetExplorer.Application")
WB.Visible := True
WB.Navigate(URL)
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10




wb.document.getElementById("uname").value := Loginname  ;ID 입력
wb.document.getElementById("login_pwd").value := Password ; 비밀번호 입력
wb.document.getElementsByTagName("INPUT")[2].Click() ; 로그인 버튼 누르기

While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10
wb.document.getElementsByTagName("A")[7].Click() ; 로그아웃 버튼 누르기




MsgBox, pause

*/


CompanyName = WESTERN LEGACY TRADING CO
;CompanyName = WESTERN

UPSLabelPrintOut()

MsgBox, pause ups



;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;History 화면에서 Tracking Number 얻기



		; A_ScreenHeight 값이 1080이라 살짝 오른쪽으로 옮기기 위해 190 더하기
		; A_ScreenWidth 값이 1920이라 화면 구석으로 옮기기 위해 1920 빼기


CompanyName = CARTERS CORNER

		Clipboard := 
		
		; 리스트의 Company Name 읽어오기
		PointOfX = %A_ScreenHeight%
		PointOfX += 435
		PointOfY = %A_ScreenWidth%
		PointOfY -= 1650
		
		MouseMove, %PointOfX%, %PointOfY%
		SendInput, #q
		Sleep 1000
		
;		PointOfX = %A_ScreenHeight%
		PointOfX -= 129
;		PointOfY = %A_ScreenWidth%
		PointOfY += 19
		
		MouseMove, %PointOfX%, %PointOfY%
		Sleep 1000
		SendInput, #q
		Sleep 1000
		MouseClick, l, PointOfX-50, PointOfY-10

		; Capture2Text 창 닫기
		IfWinExist, Capture2Text - OCR Text
			WinClose


		FindCorrectTrackingNumber := % Clipboard
;		MsgBox, % FindCorrectTrackingNumber
		
		WinActivate, LAMBS
		
		; 찾은 FindCorrectTrackingNumber 값과 CompanyName 이 같으면 마우스 마우스 오른쪽 버튼 눌러서 Tracking Number 얻기
		if(CompanyName == FindCorrectTrackingNumber){
			;MsgBox, YesFint matched
			MouseClick, r
			Loop, 33
				Send, {down}
		}
		else{
			MsgBox, NONONot matched
			Reload
		}				
	




;		if(FindCorrectTrackingNumber == Invoice_No)
;			MouseClick, l, 665, 284
;		else{
;			MsgBox, 첫째 줄의 Order Id 와 Invoice_No가 맞지 않습니다
;			Reload
;		}		
		
	
		Sleep 200





;MsgBox, %A_ScreenHeight%
;MsgBox, % aa



;return






Exitapp
 
 

Esc::


; CustomerNoteOnWeb.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1

; StaffOnlyNote.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1

; CCinfo.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfo.txt, 1
	



 Exitapp
