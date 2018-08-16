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


/*
; 이거 원래 GetActiveBrowserURL.ahk 파일 안에 있던 함수인데 이게 메인에 선언되야 com 으로 처리한 변수들의 값이 유지되어 메인에서 사용할 수 있다.
Menu, Tray, Icon, % A_WinDir "\system32\netshell.dll", 86 ; Shows a world icon in the system tray

ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass"
*/



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
	
	
	
	CompanyName := "WAREHOUSE" ;"CHANTILLY BOUTIQUE"
	
	Attention = KURT SCHOLLA ;111MITZY BURROUGHS
	Address1 = 4056 BROADWAY SUITE # 1 ;3834 CENTRAL AVE SUITE A
	Address2 = SUITE A
	ZipCode = 64111 ;91020 ;71913
	City = KANSAS CITY ;HOT SPRINGS NATIONAL PARK
	State = CA ;MO
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
	


wintitle = LAMBS -  Garment Manufacturer & Wholesale Software

F4::	
		; 만약 Please enter 'Company Name' first 경고창이 뜨면 
		; 다시 Company Name 입력하고 Sales Orders 버튼 클릭하기
		IfWinActive, ahk_class #32770
		{
			WinClose, ahk_class #32770
													
			; CompanyName 변수값 넣기
			DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
			Sleep 100
			MouseGetPos, , , , control, 1
			ControlSetText, %control%, %CompanyName%, %wintitle%
			ControlClick %control%, %wintitle%
			SendInput, {Enter}
			
			; Sales Orders 버튼 클릭 하기
			MouseClick, l, 232, 388	
		}

return






	;MsgBox, 4100, , The Number is %CurrentOrderIdNumber%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
	
	CustomerMemoOnLAMBS :=
	SalesOrderMemoONLAMBS :=
	StaffOnlyNoteVal := "ㅗㅑ"
	
	
	MsgBox, %SalesOrderMemoONLAMBS%    %StaffOnlyNoteVal%
	
	if(!CustomerMemoOnLAMBS & !SalesOrderMemoONLAMBS & !StaffOnlyNoteVal)
		MsgBox, print out
	
MsgBox, 00



	
MsgBox
CompanyName = ANCHORED SOULZ BOUTIQUE
UPSLabelPrintOut()

MsgBox, End UPSLabelPrintOut function






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





;GetInvoiceBalanceOnLAMBS()

; 79126
; cc 정보 4 개 있음

;GetInfoFromLAMBS()
;CConLAMBS(1)


;GetInfoFromFashiongo("MTR1C4178BF0A", 2)  ; 이건 결제, 배송 두 주소가 다른것
;GetInfoFromFashiongo("MTR1C26891212", 2)
;GetInfoFromFashiongo("MTR2CE7A4FB4", 2)

;OrganizingFASHIONGOCCinfo()
;GetInvoiceBalanceOnLAMBS()


;GetInfoFromLASHOROOM("OP073026207")


	MsgBox, <<Main 에서 실행>>`n`n`n`n이름 : %CompanyName%`n`n수령인 : %Attention%`n`n주소1 : %Address1%`n`n주소2 : %Address2%`n`n우편번호 : %ZipCode%`n`n주(州) : %State%`n`n도시명 : %City%`n`n전번 : %Phone%`n`n가격(Sub Total) : %SubTotal%`n`n이멜 : %Email%`n`n가격(Invoice Balance) : %InvoiceBalance%`n`n`n청구소주소 :  %BillingAdd1%`n`n청구소우편번호:  %BillingZip%`n`n`n카드번호 : %CCNumbers%`nCVV : %CVV%`nMonth : %Month%`nYear : %Year%`n`n카드번호2 : %CCNumbers2%`nCVV2 : %CVV2%`nMonth2 : %Month2%`nYear2 : %Year2%`n`n카드번호3 : %CCNumbers3%`nCVV3 : %CVV3%`nMonth3 : %Month3%`nYear3 : %Year3%`n`n카드번호4 : %CCNumbers4%`nCVV4 : %CVV4%`nMonth4 : %Month4%`nYear4 : %Year4%


loc_of_MostRecentPo = 1
OnlinePayment(loc_of_MostRecentPo)
MsgBox, online payment































/*
	`nCVV : %CVV%`nMonth : %Month%`nYear : %Year%
	`nCVV2 : %CVV2%`nMonth2 : %Month2%`nYear2 : %Year2%
	`nCVV3 : %CVV3%`nMonth3 : %Month3%`nYear3 : %Year3%
	`nCVV4 : %CVV4%`nMonth4 : %Month4%`nYear4 : %Year4%
*/	

/*
CCNumbers := "4833160135892138"
Month := "06"
Year := "19"
CVV := "180"
InvoiceBalance = 0.5


CCNumbers := "4128004084899153"
Month := "08"
Year := "20"
CVV := "575"
InvoiceBalance = 118.00
*/

/*
;GetInfoFromLAMBS()
;loc_of_MostRecentPo 값이 1이면 쇼,전화,이메일 주문인데 그냥 1넘겨봤음. 실험때는 큰 의미 없을듯
CConLAMBS(1)
;OrganizingLAMBSCCinfo(CCNumbers, ExpDate, CVV)
MsgBox, %CCNumbers%`n`n%Month%`n`n%Year%`n`n%CVV%`n`n%BillingAdd1%`n`n%BillingZip%
*/




loop, 4{
	CCNumbers%A_Index% = %A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%%A_Index%
	CVV%A_Index% = %A_Index%%A_Index%%A_Index%
	Month%A_Index% = 03
	Year%A_Index% = 2017
;	MsgBox, % CCNumbers1
;	MsgBox, % CCNumbers2

	
;	MsgBox, CCNumbers%A_Index%`n`nCVV%A_Index%`n`nMonth%A_Index%`n`nYear%A_Index%
}
; `n`n
	MsgBox, Month1 : %Month1%`n`nMonth2 : %Month2%`n`nMonth3 : %Month3%`n`nMonth4 : %Month4%
	Month2 = 02
	Year2 = 2022
	Month3 = 03
	Year3 = 2023
	Month4 = 04
	Year4 = 2024
; 카드 승인 취소 대비해서 loc_of_MostRecentPo 값을 넘겨줘야 하나?
loc_of_MostRecentPo = 1
OnlinePayment(loc_of_MostRecentPo)
MsgBox, online payment



;GetInfoFromFashiongo("MTR2CE7A4FB4", 2)
;CConFashiongo("MTR2CE7A4FB4", 2)

GetInfoFromFashiongo("MTR1C26891212", 2)
;CConFashiongo("MTR1C26891212", 2)

; 이건 결제, 배송 두 주소가 다른것
;GetInfoFromFashiongo("MTR1C4178BF0A", 2)
;CConFashiongo("MTR1C4178BF0A", 2)

OrganizingFASHIONGOCCinfo()




	; CCinfo.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
	;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
	;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfo.txt, 1
	

		MsgBox, 함수 안에서 찾은 것 CCNumbers`n%CCNumbers%`n`n`nCVV`n%CVV%`n`n`nMonth`n%Month%`n`n`nYear`n%Year%


;CConFashiongo("MTR2CE7A4FB4", 2)
*/














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


; CCinfoOnFASHIONGO.txt 내용을 CCinfoOnFASHIONGO 변수에 저장하기
FileRead, CCinfoOnFASHIONGO, %A_ScriptDir%\CreatedFiles\CCinfo.txt

; CCinfoOnFASHIONGO.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfo.txt, 1



;CustomerNoteOnWeb := GetInfoFromFashiongo("MTR1C26891212", 2)

;MsgBox, fashiongo pause

MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nState : %State%`n`nPhone : %Phone%`n`nEmail : %Email%`n`nCustomerNoteOnWeb : %CustomerNoteOnWeb%`n`nStaffOnlyNoteOnWeb : %StaffOnlyNote%`n`nCCinfoOnFASHIONGO : %CCinfoOnFASHIONGO%`n`nSubTotal : %SubTotal%`n`nBillingAdd1 : %BillingAdd1%`n`nBillingZip : %BillingZip%


;CustomerNoteOnWeb, StaffOnlyNote

MsgBox, cc




GetInfoFromLASHOROOM("OP073026207")
MsgBox, PAUSE LASHOWROOM

MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nState : %State%`n`nPhone : %Phone%`n`nEmail : %Email%`n`nCustomerNoteOnWeb : %CustomerNoteOnWeb%`n`nStaffOnlyNoteOnWeb : %StaffOnlyNote%`n`nCCinfoOnFASHIONGO : %CCinfoOnFASHIONGO%`n`nSubTotal : %SubTotal%
;WinClose, LAShowroom.com - Internet Explorer
; 마지막으로 열린 IE 창 닫힘
WinClose, ahk_class IEFrame

MsgBox, close ie??








CustomerUPSAccount := 
OriginalShippingFee :=
;OriginalShippingFee = 50

; 고객 UPS Account가 없는데(있다면 당연히 배송비는 0 되는게 맞지만)
if(!CustomerUPSAccount){
	; 배송비 가격이 없다면
	if(!OriginalShippingFee){
		; 루프를 5번 돌아라. 만약 배송비를 읽었으면 5번 이전이라도 루프를 나와라
		Loop, 5
		{
			MsgBox, no
			
			OriginalShippingFee = 1
			
			if(OriginalShippingFee)
				break
		}
	}		
}



/*
while(!OriginalShippingFee){
	MsgBox, in
}
*/
MsgBox, out




;CompanyName = THE MINT JULEP BOUTIQUE LLC
;CompanyName = GYPSY GIRLS BOUTIQUI
;CompanyName = SWEET TEXAS TREASURES
;CompanyName = LIZ AND HONEY
;CompanyName = LIME LUSH BOUTIQUE
CompanyName = ANCHORED SOULZ BOUTIQUE
;UPSLabelPrintOut()
ComparingCompanyName = ANCHORED SOULZ BOU'
ComparingCompanyName = ANCHORED SOULZ BOU
ComparingCompanyName := "8:11:18 PM"
;ComparingCompanyName = %Clipboard%
ModifiedCompanyName := "ANCHORED SOULZ  $#@%^#$&(*&)(       " ;BOU'
;ModifiedCompanyName = ANCHORED SOULZ BOUTIQUE

;ModifiedCompanyName := SubStr(CompanyName, 1, 16) ;, 19)

;ComparingCompanyName = aabb
;ModifiedCompanyName := "aa            "
;ModifiedCompanyName := "A"

; ModifiedCompanyName 에서 모든 Space(공란)를 제거합니다
;StringReplace, ModifiedCompanyName, ModifiedCompanyName, %A_Space%, , All
;StringReplace, ComparingCompanyName, ComparingCompanyName, " `t", , All

;MsgBox, %ComparingCompanyName%`n%ModifiedCompanyName%
MsgBox, %ModifiedCompanyName%

ModifiedCompanyName := Trim(ModifiedCompanyName, " `t") ;, OmitChars = " `t")
ComparingCompanyName := Trim(ComparingCompanyName, " `t") ;, OmitChars = " `t")
;ModifiedCompanyName := Trim(ModifiedCompanyName, " ") ;, OmitChars = " `t")

MsgBox, %ModifiedCompanyName%
MsgBox, % ComparingCompanyName

;if(RegExMatch(ComparingCompanyName, %ModifiedCompanyName%))
;IfInString, ComparingCompanyName, %ModifiedCompanyName%
;if(RegExMatch(ComparingCompanyName, "imU)\d\d:\d\d.*AM|PM"))
if(RegExMatch(ComparingCompanyName, "imU)\d*:\d*.*AM|PM"))
{
	MsgBox, found
}

MsgBox, pa




UPSLabelPrintOut()

MsgBox, FUNCTION OUT
*/


/*
;WinActivate, UPS

;asdfasdfasdfasdf("MTR2CE7A4FB4", 2)
;MsgBox, asdfasdfasdfasdf
*/

/*
GetInfoFromFashiongo("MTR2CE7A4FB4", 2)

; 현재 존재하는 IE창 접속 하기 위해 함수 호출
;WBGet(WinTitle="ahk_class IEFrame", Svr#=1)             ;// based on ComObjQuery docs
IEGet(name="")

aa := CConFashiongo("MTR2CE7A4FB4", 2)
OrganizingCCinfoOnFASHIONGO()

MsgBox aa is %aa%
*/













;Sleep 500
i=1
j=abc3
	Loop, 5
	{
		
		ComparingCompanyName = abc%a_index%
		MsgBox, Company name %a_index% is `n%ComparingCompanyName%

;		ComparingCompanyName := SortedCompanyName1
;		MsgBox, Company name 1 is `n%ComparingCompanyName%
;		Clipboard := % ComparingCompanyName
;		MsgBox, % CompanyName


		; 읽어온 회사명(ComparingCompanyName)값이 CompanyName(찾고 있는 회사이름)과 같으면 이 페이지에 찾고 있는 회사명이 있다는 얘기
		; 그러니까 현재 페이지의 처음 회사명부터 차례대로 하나씩 다시 찾아보기
		IfEqual, ComparingCompanyName, %j%
		{
			
;		if(ComparingCompanyName == CompanyName){
			MsgBox, matched

			PointOfX = %A_ScreenWidth%
			PointOfY = %A_ScreenHeight%
			PointOfX -= 605
		}
		MsgBox, out?
		i += 1
	}




	; 현재 페이지에 있는 모든 값을 읽기위한 세팅
	; PointOftoY 값을 회사명 끝까지 내렸음
	PointOfX = %A_ScreenWidth%
	PointOfY = %A_ScreenHeight%
	PointOfX -= 600
	PointOfY -= 808
	
	PointOftoX = %A_ScreenWidth%
	PointOftoY = %A_ScreenHeight%
	PointOftoX -= 400
	PointOftoY -= 530
	
	; 지정된 위치(현재 페이지)에서 값 얻기
	MouseMove, %PointOfX%, %PointOfY%
	SendInput, #q
	Sleep 200
	MouseMove, %PointOftoX%, %PointOftoY%
	Sleep 200
	SendInput, #q
	Sleep 200
				
	; Capture2Text 창 닫기
	WinWaitActive, Capture2Text - OCR Text
	WinClose
	
;	IfWinExist, Capture2Text - OCR Text
;		WinClose


	AllCompanyNamesOnCurrentPage := % Clipboard
	Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨

	MsgBox, % AllCompanyNamesOnCurrentPage
	
	
	; AllCompanyNamesOnCurrentPage 값에 개행문자(새 줄)이 나올때마다 나눠서 SortedCompanyName 에 저장
	StringSplit, SortedCompanyName, AllCompanyNamesOnCurrentPage, `n, %A_Space%
	;StringSplit, Wts_of_Boxes, Invoice_Wts, %A_Space%, `,|`.  ; 점이나 콤마는 제외합니다.

	; SortedCompanyName 에 들어있는 값 갯수만큼만 루프 돌려서
	; 
	Loop, %SortedCompanyName0%
	{
		
		ComparingCompanyName := SortedCompanyName%a_index%
		MsgBox, Company name %a_index% is `n%ComparingCompanyName%

;		ComparingCompanyName := SortedCompanyName1
;		MsgBox, Company name 1 is `n%ComparingCompanyName%
		Clipboard := % ComparingCompanyName
		MsgBox, % CompanyName


		; 읽어온 회사명(ComparingCompanyName)값이 CompanyName(찾고 있는 회사이름)과 같으면 이 페이지에 찾고 있는 회사명이 있다는 얘기
		; 그러니까 현재 페이지의 처음 회사명부터 차례대로 하나씩 다시 찾아보기
;		IfEqual, ComparingCompanyName, %CompanyName%
;		{
			
		if(ComparingCompanyName == CompanyName){
			MsgBox, matched

			PointOfX = %A_ScreenWidth%
			PointOfY = %A_ScreenHeight%
			PointOfX -= 605
			PointOfY -= 808
			
			PointOftoX = %A_ScreenWidth%
			PointOftoY = %A_ScreenHeight%
			PointOftoX -= 400
			PointOftoY -= 792
			
			
			; SortedCompanyName 에 들어있는 값 갯수만큼만 루프 돌려서
			Loop, %SortedCompanyName0%
			{

				; 지정된 위치에서 값 얻기
				MouseMove, %PointOfX%, %PointOfY%
				SendInput, #q
				Sleep 200
				MouseMove, %PointOftoX%, %PointOftoY%
				Sleep 200
				SendInput, #q
				Sleep 200
				
				; 그 다음 회사명을 찾기 위해 Y값들 재설정 하기
				PointOfY += 16
				PointOftoY := PointOfY + 15
				
				Sleep 1000
				WinClose, Capture2Text - OCR Text
				
				
				
				HavingBeenReadIndividualCompanyName := % Clipboard
							
							
				; 읽어온 회사명(HavingBeenReadIndividualCompanyName)값이 CompanyName과 같으면 마우스 오른쪽 버튼 눌러서 Tracking Number 얻기
				if(HavingBeenReadIndividualCompanyName == CompanyName){
					Clipboard :=
					PointOfX += 10
					PointOfY += 10
											
					MouseClick, r, %PointOfX%, %PointOfY%
					Sleep 200
					Loop, 33{
						SendInput, {Down}
			;			Sleep 100
					}
							
					SendInput, {Enter}
					Sleep 700 ;클립보드 값을 사용하기 위해서는 최소 0.7초는 기다려야됨
					

						
					TrackingNumber := % clipboard
					
					MsgBox, % clipboard

					MsgBox, 읽어온 회사명:%ReadVal%`n`n송장번호: %TrackingNumber%
					
					; 읽어온 회사명과 같은 것을 찾았으면 루프문을 아예 탈출해야 되는데 뭘로 할까 goto를 써야하나
					break
				}				
				
			}
		}








		else{
			;찾는 회사명이 없으면 스크롤 다운 해서 다음 페이지에서 다시 찾아봐야 됨
		}

	MsgBox, loop out?
	}		
















































/*
		GetInfoFromFashiongo("MTR47CEFB0C", 2)
		CConFashiongo("MTR47CEFB0C", 2)
		
		MsgBox, CompanyName : %CompanyName%`n`nAttention : %Attention%`n`nAddress1 : %Address1%`n`nAddress2 : %Address2%`n`nZipCode : %ZipCode%`n`nCity : %City%`n`nState : %State%`n`nPhone : %Phone%`n`nEmail : %Email%

		
		IfWinExist, , Internet Explorer
			WinClose
*/





/*
Loginname = customer3
Password = Jo123456789
URL = http://vendoradmin2.fashiongo.net/OrderDetails.aspx?po=MTR47CEFB0C

WB := ComObjCreate("InternetExplorer.Application")
WB.Visible := true
WB.Navigate(URL)
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10
   
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
	wb.document.getElementById("tbUserID").value := Loginname  ;ID 입력
	wb.document.getElementById("tbPassword").value := Password ; 비밀번호 입력
	wb.document.getElementsByTagName("A")[0].Click() ; 로그인 버튼 누르기

	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
	   Sleep, 10
   
;	MsgBox, found login
}


; html source 얻기
htmlSourcecode := WB.Document.All[0].outerhtml

; 패션고 페이지에서 고객정보 읽어오기
FindInfoInFASHIONGO(htmlSourcecode)

*/



Exitapp

F7::
Reload
	
Esc::
 ; UPS 창 항상위에 설정 해제
 Winset, AlwaysOnTop, Off, UPS WorldShip - Remote Workstation
 
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
 Reload