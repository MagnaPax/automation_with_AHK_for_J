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


global TrackingNumber, UserName, Decline1st, Decline2nd, Decline3rd, InvoiceBalance

global InvoiceMemoOnLAMBS, CustomerMemoOnLAMBS, CustomerNoteOnWeb, StaffOnlyNote, Invoice_No, FGServer, wb, Paymentwb ;, CCinfoOnFASHIONGO


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
	




; Capture2Text 실행되고 있는 지 확인 후 실행 안 되고 있으면 실행
Process, Exist, Capture2Text.exe ; check to see if Capture2Text.exe is running
{		
	If ! errorLevel
	{
		;~ MsgBox, Capture2Text_64bit is not running, If Ok Button Clicked, the Application will be Run.
		IfExist, %A_ScriptDir%\Capture2Text_64bit\Capture2Text.exe
			Run, %A_ScriptDir%\Capture2Text_64bit\Capture2Text.exe
		;~ Return
	}
}
	
	


;GUI Backgroud
;Gui, Show, w327 h137, JODIFL
Gui, Show, w350 h255, JODIFL

;Input invoice Number
;Gui, Add, Text, x22 y19 w70 h20 , Invoice No.
Gui, Add, Text, x22 y19 Cred , Invoice No.
Gui, Add, Edit, x92 y19 w100 h20 vInvoice_No,  ;70829 ;67629 ;Puerto Rico Order 72253

;Input Wts of boxes
;Gui, Add, Text, x22 y49 w70 h20 , Wts
Gui, Add, Text, x22 y49 CBlue , Wts
Gui, Add, Edit, x92 y49 w100 h20 vInvoice_Wts,  ;11, 22, 33

;Input Quantity of Items
Gui, Add, Text, x22 y79 w70 h20 , # of Items
Gui, Add, Edit, x92 y79 w100 h20 vH1,
Gui, Add, UpDown, x172 y79 w20 vQty_of_Items, 1

;Customer's UPS Account
Gui, Add, Text, x20 y120 w85 h40 +Center, Customer's UPS Account
Gui, Add, Edit, x22 y151 w85 h20 vCustomerUPSAccount, ;X870y4

;Apply Credit 체크박스
Gui, Add, CheckBox, x22 y180 w90 h30 vApplyCredit +Center, Apply Credit

;Consolidation 체크박스
Gui, Add, CheckBox, x22 y210 w95 h30 vConsolidation +Center, Consolidation


;2nd Month 체크박스
Gui, Add, CheckBox, x140 y115 w65 h30 v2ndMonth +Center, 2nd Month

;3rd Month 체크박스
Gui, Add, CheckBox, x140 y145 w65 h30 v3rdMonth +Center, 3rd Month

;Next Month 체크박스
Gui, Add, CheckBox, x140 y175 w70 h30 vNextMonth +Center, Next Month


;Delivery
Gui, Add, Text, x135 y210 w85 h40 +Center, Delivery
Gui, Add, Edit, x137 y225 w85 h20 vDelivery, ;X870y4


;No Declare Request 체크박스
Gui, Add, CheckBox, x245 y125 w95 h50 vNoDeclare +Center, UPS`nDon't Declare`na Value

;FashionGo Server Choosing
Gui, Add, Text, x250 y200 w70 h20 , FG URL #
Gui, Add, Edit, x252 y220 w50 h20 -Tabstop vH2,
Gui, Add, UpDown, x250 y250 w20 vFGServer, 2



;엔터 버튼
Gui, Add, Button, x225 y19 w100 h80 +default gClick_btn, Enter



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
	
	; 혹시 모르니 메모리에 있는 정보 초기화 하기
	Clipboard :=
	
	; LAMBS예 정보 입력
	No_of_Boxes := 1stCommonLAMBSProcessing(Invoice_No, Invoice_Wts, No_of_Boxes, ApplyCredit, Consolidation, CustomerUPSAccount, NextMonth, 2ndMonth, 3rdMonth, Qty_of_Items)



; CustomerNoteOnWeb.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1

; StaffOnlyNote.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1

; CCinfo.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfo.txt, 1

; URLofVirtualPOSTerminal.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt, 1








		; 만약 Consolidation 이 체크 됐으면 ConsolidationProcessing 함수 호출
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;조디플 홈피는 주문 시 처리 만들어야 되고;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		if(Consolidation)
			ConsolidationProcessing(ApplyCredit, Invoice_Memo, Invoice_No, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L, Delivery, FGServer)
		
		if(Delivery)
			ConsolidationProcessing(ApplyCredit, Invoice_Memo, Invoice_No, FoundPos, F_arr, L_arr, W_arr, i, j, k, lv_F, lv_W, lv_L, Delivery, FGServer)















;	MsgBox, Set the searching status of FASHIONGO

	Invoice_Memo := % InvoiceMemoOnLAMBS
	;Invoice_Memo = , Sales #41668/PO #MTR1B08BD00EA, Sales #42218/PO #MTR1B1FF6DECF TRACKING# SENT EMAIL -EUNICE 6/5/17 LVM&LEM 6/8/17 KRISLEN // LVM 06/12/2017, SENT EMAIL-JAZMINE//   Customer was contacted through phone call + email and no response was received. Order cancelled due to lack of payment. KRISLEN** 6/14/2017. TRACKING #
	
;	MsgBox, % Invoice_Memo
	
;	sleep 1000
	
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
		;MsgBox, 쇼,전화,이메일 주문입니다. ;FoundPos: `n%FoundPos%
		
		;웹에서 온 주문이 아니기 때문에 LAMBS에서 정보 읽고 저장하는 함수 호출
		GetInfoFromLAMBS()
;		MsgBox, Out of GetInfoFromLAMBS
		
		;return
	}
	else if(loc_of_MostRecentPo = lv_F){
		GetInfoFromFashiongo(F_PO, FGServer)
		OrganizingFASHIONGOCCinfo()
	}

	else if(loc_of_MostRecentPo = lv_W)
		;GetInfoFromJODIFL_WEB(W_PO)
		GetInfoFromLAMBS()
	else
		GetInfoFromLASHOROOM(L_PO)
	
	
	

;		MsgBox, % CustomerNoteOnWeb






	; CustomerNoteOnWeb.txt 내용을 CustomerNoteOnWeb 변수에 저장하기
	FileRead, CustomerNoteOnWeb, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt

	; CustomerNoteOnWeb.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
	;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1
	FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1


	; StaffOnlyNote.txt 내용을 StaffOnlyNote 변수에 저장하기
	FileRead, StaffOnlyNote, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt

	; StaffOnlyNote.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
	;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1
	FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1


		
	
	
	
	

	
	
	
	
	
	
	
	
	
	
	
	
	

	
	
	
	
	
	
	
	
	
	; LAMBS의 인보이스메모, 커스토머메모, 웹의 커스토머메모 띄우기
	MsgBox, -Invoice Memo On LAMBS-`n`n`n%InvoiceMemoOnLAMBS%`n`n`n`n`n`n`n-Customer Memo On LAMBS-`n`n`n%CustomerMemoOnLAMBS%`n`n`n`n`n`n`n-Customer Notes on Web-`n`n`n%CustomerNoteOnWeb%`n`n`n`n`n`n`n-Staff Only Note on Web-`n`n`n%StaffOnlyNote%
	
	

	
	
	
	
	
	

	
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
	RoundedShippingFee := UPS_InputAdd(NextMonth, 3rdMonth, 2ndMonth, CustomerUPSAccount, Invoice_No, No_of_Boxes, Invoice_Wts, NoDeclare)
;	MsgBox, % RoundedShippingFee


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
/*
	; 주소, 전화번호 등 고객 정보 얻기 위해 열었던 IE 창 닫기
	; 마지막으로 열린 IE 창 닫힘
	IfWinExist, ahk_class IEFrame
		WinClose
	wb :=
*/
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


	;화면 초기화
	Start()
	Start_Invoice_1()
	
	; EST/Actual FC에 배송비 입력하기	
	DragAndPast(810, 336, 730, 336, RoundedShippingFee)
	SendInput, {Tab}

	; 배송비 입력됐으니 저장하기
	SendInput, {F8}

	;security 버튼 찾아서 풀어주기
	FindSecurityButtonAndClickItThenInputNumber1()
	Sleep 500
	
	; 만약 ApplyCredit 가 체크 됐으면 ApplyCredit함수 호출
	if(ApplyCredit){
		
		ApplyCreditFunction(Invoice_No)
		Sleep 500
		
		; Credit 적용된 이후에는 값이 바뀌었을 테니 Invoice Balance 값 다시 저장하기
		InvoiceBalance := ClickAndCtrlAll(867, 591)
		SendInput, {F8}
	}
	
	
	; Invoice Balance 값 얻기
	InvoiceBalance := ClickAndCtrlAll(867, 591)
		

/*
	; 쇼, 전화, 이메일 주문이라면 LAMBS에서 CC 창 열기
	if(loc_of_MostRecentPo = 1){
		CConLAMBS(loc_of_MostRecentPo)
;		MsgBox, cc WINDOW oUT
	}
*/

	; 각 주문별 CC함수 호출
	if(loc_of_MostRecentPo = 1)
		CConLAMBS(loc_of_MostRecentPo)
	else if(loc_of_MostRecentPo = lv_F)
	{
		;CConFashiongo(PO_F, FGServer)
		;OrganizingFASHIONGOCCinfo()
		
		; GetInfoFromFashiongo ; 함수 호출할 때 이미 cc정보 읽어서 여기서 따로 다른 함수를 호출 할 필요 없음
		; 하지만 Apply Credit이 만약 적용됐으면 바뀐 Invoice Balance 가격을 읽어와야 되기 때문에 InvoiceBalance 값 읽는것만 실행
		GetInvoiceBalanceOnLAMBS()

	}
	else if(loc_of_MostRecentPo = lv_W)
		CConLAMBS(loc_of_MostRecentPo) ; jodifl.com 주문은 쇼,전화,이메일 주문과 처리법이 동일해서 LAMBS 에서 CC 정보 얻는다
	else{
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; CConLASHOROOM 함수는 구현해야 됨 ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		CConLASHOROOM()
		GetInvoiceBalanceOnLAMBS()
	}
		



	MsgBox, <<Main 에서 실행>>`n`n`n`n이름 : %CompanyName%`n`n수령인 : %Attention%`n`n주소1 : %Address1%`n`n주소2 : %Address2%`n`n우편번호 : %ZipCode%`n`n주(州) : %State%`n`n도시명 : %City%`n`n전번 : %Phone%`n`n가격(Sub Total) : %SubTotal%`n`n이멜 : %Email%`n`n가격(Invoice Balance) : %InvoiceBalance%`n`n`n청구소주소 :  %BillingAdd1%`n`n청구소우편번호:  %BillingZip%`n`n`n카드번호 : %CCNumbers%`nCVV : %CVV%`nMonth : %Month%`nYear : %Year%`n`n카드번호2 : %CCNumbers2%`nCVV2 : %CVV2%`nMonth2 : %Month2%`nYear2 : %Year2%`n`n카드번호3 : %CCNumbers3%`nCVV3 : %CVV3%`nMonth3 : %Month3%`nYear3 : %Year3%`n`n카드번호4 : %CCNumbers4%`nCVV4 : %CVV4%`nMonth4 : %Month4%`nYear4 : %Year4%
	


	; 자동 결제하기 위한 함수 호출
	; 카드 승인 취소 대비해서 loc_of_MostRecentPo 값을 넘겨줘야 하나?
	OnlinePayment(loc_of_MostRecentPo)



	; CCinfoOnFASHIONGO.txt 내용을 CCinfoOnFASHIONGO 변수에 저장하기
;	FileRead, CCinfoOnFASHIONGO, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt

	; CCinfo.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
	;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
	FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfo.txt, 1
	





	; 카드 결제가 제대로 됐는지 묻고 Yes 눌렀으면 UPS Label 인쇄 함수 호출
	MsgBox, 4100, UPS Label Print out, Does Credit Card Go Passed?
	
	IfMsgBox, Yes
	{
		UPSLabelPrintOut()
		
		;LAMBS의 Invoice Memo에 Tracking Number 적어넣는 함수호출
;		WrapUp1st(Consolidation, loc_of_MostRecentPo, lv_F, lv_W)
		WrapUp1st(Consolidation, loc_of_MostRecentPo, Invoice_Memo, Invoice_No, Delivery)
		
;		Reload
	
	}
	else IfMsgBox, No
	{
		; 쇼,전화,이메일 주문이면 이미 CConLAMBS 열어본 상태이므로 그냥 DeclineProcessing 호출
		if(loc_of_MostRecentPo = 1){
			;Decline 이메일 보내는 함수 호출
			;MsgBox, NO~~~
			DeclineProcessing()
		}
		else ; CConLAMBS 함수 호출 후 DeclineProcessing 호출하는
			CallCConLAMBSAndAskWhetherCCwentOrNot(Consolidation, loc_of_MostRecentPo, Invoice_Memo, Invoice_No)
	}
	
	; 쇼,전화,이메일 주문이면 그냥 리로드하기
	if(loc_of_MostRecentPo = 1)
		Reload
	
	
	; 웹 주문들의 마지막 기본 페이지 열기
	OpenAllBaseWebPageOfOrders(i, j, k, F_arr, W_arr, L_arr, FGServer)


	; 쇼,전화,이메일 주문이 아니면 웹에 값을 입력하기 위해 대기하기
	MsgBox, 4100, Input values at Web, Ready to Reload


	; 주소, 전화번호 등 고객 정보 얻기 위해 열었던 IE 창 닫기
	; 마지막으로 열린 IE 창 닫힘
	IfWinExist, ahk_class IEFrame
		WinClose

	Reload

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

	; 2ND Month, 3RD Month UPS처리 해줘야 됨
	; consolidation 처리 해줘야 됨
	
	; decline 됐을 때 이메일 보내는 것 완전히 끝내지 않았음
	; 주문이 추가된 경우에는 나중의 PO가 더 오래된 것일 수 있다. 백오더가 나중에 더해질 수도 있어서 72167
	
	; DeclineProcessing 함수 호출 시작할 때 언어 설정을 영어로 바꾸는 것 해야지 오류 안 나지 않을까
	
	
	
	
	



!F9::
SendInput, % Invoice_No
return

!F10::
SendInput, % TrackingNumber
return

!F11::
SendInput, % RoundedShippingFee
return


F7::
; CustomerNoteOnWeb.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1

; StaffOnlyNote.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1

; CCinfo.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfo.txt, 1

; URLofVirtualPOSTerminal.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt, 1

	
Reload


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

; URLofVirtualPOSTerminal.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt, 1
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\URLofVirtualPOSTerminal.txt, 1

	
 Exitapp
 Reload