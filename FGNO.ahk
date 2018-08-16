#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include GetActiveBrowserURL.ahk
#Include FindTextFunctionONLY.ahk



global CurrentOrderIdNumber, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, CustomerNoteOnWebVal, StaffOnlyNoteVal, CompanyName, PendingOrderStatus, AlreadyProcessedItem, CurrentPONumber, ClickLOC, AdditionalInfo, driver, sURL, PO_Number, Order_ID_Only

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




; 이거 원래 GetActiveBrowserURL.ahk 파일 안에 있던 함수인데 이게 메인에 선언되야 com 으로 처리한 변수들의 값이 유지되어 메인에서 사용할 수 있다.
Menu, Tray, Icon, % A_WinDir "\system32\netshell.dll", 86 ; Shows a world icon in the system tray

ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
LegacyBrowsers := "IEFrame,OperaWindowClass"




;GUI Backgroud
Gui, Show, w350 h150, NewOrdersProcessing, AlwaysOnTop Window
WinSet, AlwaysOnTop, On, NewOrdersProcessing

;Input Start Order Id
Gui, Add, Text, x22 y21 Cred , Start Order ID
Gui, Add, Edit, x92 y19 w100 h20 vStartOrderId,  ;53493 ;49998 ;49993

;Input End Order Id
Gui, Add, Text, x22 y51 CBlue , End Order ID
Gui, Add, Edit, x92 y49 w100 h20 vEndOrderId,  ;11, 22, 33

;PO NO.
Gui, Add, Text, x22 y83 CGreen , PO NO.`n(For Urgent Order)
Gui, Add, Edit, x92 y79 w100 h20 vPO_Number,  ; MTR1D55EDC764

;ORDER ID
Gui, Add, Text, x22 y115 CBlack , Order ID Only`n(Skip Web Processing)
Gui, Add, Edit, x92 y109 w100 h20 vOrder_ID_Only, ; 54250

/*
;FashionGo Server Choosing
Gui, Add, Text, x22 y79 w70 h20  , FG URL #
Gui, Add, Edit, x92 y79 w100 h20 vH1 -Tabstop vH2,
Gui, Add, UpDown, x172 y79 w20 vFGServer, 2
*/


;엔터 버튼
Gui, Add, Button, x225 y19 w100 h110 +default gClick_btn, Enter



;GUI시작 시 포커스를 Invoice_No 입력칸에 위치
GuiControl, Focus, StartOrderId


return



Click_btn:


	; 이전 IE 창 열려있으면 자꾸 이런저런 에러 나서 아예 IE 창들을 모두 그룹으로 묶고 한꺼번에 닫고 시작
	;~ GroupAdd,ExplorerGroup, ahk_class CabinetWClass
	GroupAdd,ExplorerGroup, ahk_class IEFrame
	WinClose,ahk_group ExplorerGroup


	; 혹시 창을 열어놓은 상태로 중단했을 수도 있으니까 창 닫기
	IfWinExist, Transfer from Sales Order
		WinClose
	IfWinExist, Customer Order +Zoom In
		WinClose
	IfWinExist, Accounts Summary
		WinClose
	IfWinExist, Style-WIP (P990202)
		WinClose
	IfWinExist, Style-Cust Order Details (P990203)
		WinClose


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; To use the values which input on GUI
	Gui Submit, nohide
	GUI, Destroy


	; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
	BlockInput, Mouse

	
	; 시작 번호값 CurrentOrderIdNumber 변수에 넣기
	CurrentOrderIdNumber = %StartOrderId%
	
	; 마지막 번호 값 ModifiedEndOrderIDNumber 변수에 넣고 값에 1 더해주기
	ModifiedEndOrderIDNumber := % EndOrderId
	ModifiedEndOrderIDNumber++





	Loop
	{
		
		; ####################################################################################################################################################################################
		; Capture2Text 실행되고 있는지 감지 후 실행되지 않고 있으면 처리하기
		; https://www.google.com/search?q=autohotkey+detect+if+program+running&rlz=1C1CHFX_enUS656US656&oq=autohotkey+detect+running+p&aqs=chrome.2.69i57j0l2.15446j0j8&sourceid=chrome&ie=UTF-8
		; https://autohotkey.com/board/topic/33659-if-program-is-not-running-then-start/
		; ####################################################################################################################################################################################
		

		; 만약 LAMBS 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		IfWinNotExist, LAMBS
		{
			MsgBox, 262144, No LAMBS Warning, PLEASE RUN LAMBS
			continue
		}	
		

		; 혹시 창을 열어놓은 상태로 중단했을 수도 있으니까 창 닫기
		IfWinExist, Transfer from Sales Order
			WinClose
		IfWinExist, Customer Order +Zoom In
			WinClose
		IfWinExist, Accounts Summary
			WinClose



		; 마지막 번호에 도달하면 끝내기
		IfEqual, CurrentOrderIdNumber, %ModifiedEndOrderIDNumber%
			Reload

		Start()

		wintitle = LAMBS -  Garment Manufacturer & Wholesale Software
		
		; 메모리에 있는 값, PendingOrderStatus 값, AlreadyProcessedItem 값 초기화 하기
		Clipboard :=
		PendingOrderStatus :=
		AlreadyProcessedItem :=
		CurrentPONumber :=
		ClickLOC :=
		


		; Create Sales Orders Small 
		OpenCreateSalesOrdersSmallTab()
		Sleep 50
		
		; Customer PO 1 클릭하기
		Click, 59, 262
		
		;New & Clear 버튼 클릭
		MouseClick, l, 60, 125
		sleep 500

		
		;Hide All 클릭해서 메뉴 바 없애기
	;	ClickAtThePoint(213, 65)



		; gui에서 Order_ID_Only 값 입력했으면 CurrentOrderIdNumber 변수에 Order_ID_Only 값을 넣기. 이렇게 되면 코드 크게 바꾸지 않아도 되기 때문
		if(Order_ID_Only)
			CurrentOrderIdNumber := Order_ID_Only
		

		; Order ID 입력칸으로 이동 후 CurrentOrderIdNumber 변수값 넣기
		DllCall("SetCursorPos", int, 181-8, int, 190-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
	;	MouseMove, 181, 190
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlSetText, %control%, %CurrentOrderIdNumber%, %wintitle%
		ControlClick %control%, %wintitle%
		SendInput, {Enter}


		Sleep 1000


		; gui에서 PO 번호 입력했으면 함수 호출해서 처리
		if(PO_Number)
			ProcessWhenPoNumberInputted(PO_Number)
		
		
		; Company Name 얻기
		DllCall("SetCursorPos", int, 87-8, int, 375-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
	;	MouseMove, 87, 375
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, CompanyName, %control%, %wintitle%
		
/*		
		; 혹시 어떤 오류 때문에 Order ID를 입력 못해서 PO값이 없으면 다시 처음부터 시작해서 값 입력하기
		if(!CompanyName){
			MsgBox, 262144, No Company Name Warning, THERE IS NO COMPANY NAME, PEASE CHECK IT AND THEN CLICK THE BUTTON
		}
*/
		
		; Customer 밑에 있는 메모 POSourceOrMemo 변수에 저장하기
		DllCall("SetCursorPos", int, 83-8, int, 406-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
	;	MouseMove, 83, 406
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, POSourceOrMemo, %control%, %wintitle%
		POSourceOrMemo := RegExReplace(POSourceOrMemo, "imU)FASHIONGO", "")
		POSourceOrMemo := RegExReplace(POSourceOrMemo, "imU)LASHOWROOM", "")
		POSourceOrMemo := RegExReplace(POSourceOrMemo, "imU)WEB", "")

		
		; Customer Memo 값 CustomerMemoOnLAMBS 변수에 저장하기
		DllCall("SetCursorPos", int, 67-8, int, 466-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).	
	;	MouseMove, 67, 466
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, CustomerMemoOnLAMBS, %control%, %wintitle%
		StringUpper, CustomerMemoOnLAMBS, CustomerMemoOnLAMBS ; Staff only notes 대문자로 바꾸기


		
		; Sales Order Memo 값 SalesOrderMemoONLAMBS 변수에 저장하기
		DllCall("SetCursorPos", int, 69-8, int, 563-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
	;	MouseMove, 69, 563
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, SalesOrderMemoONLAMBS, %control%, %wintitle%
		SalesOrderMemoONLAMBS := RegExReplace(SalesOrderMemoONLAMBS, "imU)Handling\sFee:\s.0.00", "")
		StringUpper, SalesOrderMemoONLAMBS, SalesOrderMemoONLAMBS ; Staff only notes 대문자로 바꾸기

		
		; Shipping Add 의 State 값 읽기
		DllCall("SetCursorPos", int, 736-8, int, 591-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, State, %control%, %wintitle%
			
		
		; PO 번호 CurrentPONumber 변수에 저장하기	
		DllCall("SetCursorPos", int, 994-8, int, 378-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
	;	MouseMove, 994, 378
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, CurrentPONumber, %control%, %wintitle%




		; 혹시 어떤 오류 때문에 Order ID를 입력 못해서 PO값이 없으면 다시 처음부터 시작해서 값 입력하기
		if(!CompanyName){
			MsgBox, 262144, No Company Name Warning, THERE IS NO COMPANY NAME, PEASE CHECK IT AND THEN CLICK THE BUTTON
		}


		; 혹시 어떤 오류 때문에 Order ID를 입력 못해서 PO값이 없으면 다시 처음부터 시작해서 값 입력하기
		if(!CurrentPONumber)
			continue




		; POSourceOrMemo 값에 TIFFANY 가 들어있으면 TIFFANY ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
		IfInString, POSourceOrMemo, TIFFANY
		{
			;MsgBox, , , IT'S TIFFANY ONLY ORDER, MOVE TO NEXT ORDER`n`n`nTHIS WINDOW WILL BE CLOSED IN 5 SECONDS, 5
			MsgBox, , , IT'S TIFFANY ONLY ORDER, MOVE TO NEXT ORDER
			
			; FG 페이지는 정리해줘야 되니까 함수 열어서 처리하기
			;~ OpenFGforNewFGProcessing_CHROME()
			OpenFGforNewFGProcessing_UsingIE()
			
			CurrentOrderIdNumber++
			continue
		}

		; POSourceOrMemo 값에 JONATHAN 가 들어있으면 JONATHAN ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
		IfInString, POSourceOrMemo, JONATHAN
		{
			MsgBox, , , IT'S JONATHAN ONLY ORDER, MOVE TO NEXT ORDER
			
			; FG 페이지는 정리해줘야 되니까 함수 열어서 처리하기
			;~ OpenFGforNewFGProcessing_CHROME()
			OpenFGforNewFGProcessing_UsingIE()			
			
			CurrentOrderIdNumber++
			continue
		}

		; POSourceOrMemo 값에 ROONY 가 들어있으면 JONATHAN ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
		IfInString, POSourceOrMemo, ROONY
		{
			MsgBox, , , IT'S ROONY ONLY ORDER, MOVE TO NEXT ORDER
			
			; FG 페이지는 정리해줘야 되니까 함수 열어서 처리하기
			;~ OpenFGforNewFGProcessing_CHROME()
			OpenFGforNewFGProcessing_UsingIE()			
			
			CurrentOrderIdNumber++
			continue
		}


		; POSourceOrMemo 값에 CONNIE 가 들어있으면 CONNIE ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
		IfInString, POSourceOrMemo, CONNIE
		{
			MsgBox, , , IT'S CONNIE ONLY ORDER, MOVE TO NEXT ORDER
			
			; FG 페이지는 정리해줘야 되니까 함수 열어서 처리하기
			;~ OpenFGforNewFGProcessing_CHROME()
			OpenFGforNewFGProcessing_UsingIE()			
			
			CurrentOrderIdNumber++
			continue
		}



		; Order_ID_Only 값이 없으면, 즉 gui에서 Order_ID_Only 값을 입력하지 않았으면 FG 페이지 열어서 처리하기
		if(!Order_ID_Only){
			; 해당 PO의 패션고 페이지 열고 여러가지 처리하기
			OpenFGforNewFGProcessing_UsingIE()
			
			;~ OpenFGforNewFGProcessing_CHROME()
		}
		

		; CustomerNoteOnWebVal.txt 내용을 CustomerNoteOnWebVal 변수에 저장하기
		FileRead, CustomerMemoOnLAMBS, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt

		; StaffOnlyNoteVal.txt 내용을 StaffOnlyNoteVal 변수에 저장하기
		FileRead, StaffOnlyNoteVal, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt

		; AdditionalInfo.txt 내용을 AdditionalInfo 변수에 저장하기
		FileRead, AdditionalInfo, %A_ScriptDir%\CreatedFiles\AdditionalInfo.txt


		; CustomerNoteOnWebVal.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
		;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
		FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1

		; StaffOnlyNoteVal.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
		;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
		FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\StaffOnlyNote.txt, 1

		; AdditionalInfo.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
		;FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CCinfoOnFASHIONGO.txt, 1
		FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\AdditionalInfo.txt, 1


		
		;~ MsgBox, CustomerNoteOnWebVal : %CustomerNoteOnWebVal%`n`n`nStaffOnlyNoteVal : %StaffOnlyNoteVal%`n`n`nAdditionalInfo : %AdditionalInfo%
		
		



		; TheError 값이 1이면 웹 주문 페이지 상태가 컨펌으로 바뀌지 않은 것이므로 처음으로 돌아가 회돌이 다시 시작하기
		IfEqual, TheError, 1
			continue
		

		; AlreadyProcessedItem 값이 1이면 이미 처리된 PO 번호이므로 처음으로 돌아가 회돌이 다시 시작하기
		IfEqual, AlreadyProcessedItem, 1
		{
			WinClose, ahk_exe iexplore.exe
			;~ WinClose, ahk_class Chrome_WidgetWin_1
			CurrentOrderIdNumber++
			continue
		}
		

		
/*
		; 패션고에 고객 메모가 있으면 띄우기
		; 고객이 남긴 메모 중 혹시 이해가 안 가는 것이 있으면 여쭙기 위해
		if(CustomerNoteOnWebVal)
			MsgBox, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`n**Customer Note On FG**`n`n`n`n%CustomerNoteOnWebVal%
		
		if(StaffOnlyNoteVal)
			MsgBox, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`n**Staff Only Note On FG**`n`n`n`n%StaffOnlyNoteVal%
*/		
		
		; Account Summary 띄워서 Pending Order 있는지 확인 하기
		AccountSummayrProcessingONCreateSalesOrdersSmallTab()
		
		
		
		; 만약 펜딩값이 있었으면 AccountSummayrProcessingONCreateSalesOrdersSmallTab 함수 안에서 Create Invoice 창으로 넘어가는 처리 하고
		; PendingOrderStatus 변수에 값을 넣어서 펜딩값이 있어서 처리해줬다는 표시를 해줬음
		
		; 그 후처리 하기
		

		;####################
		;펜딩 값이 있었을 경우
		;####################
		
		if(PendingOrderStatus){
				
			; 패션고 페이지 활성화 시켜서 스크롤 다운 한 번 해주기
			; 이러면 주문 어떻게 했는지 보고 판단할 때 도움되니까
;			WinActivate, ahk_exe iexplore.exe
			;~ WinActivate, ahk_class Chrome_WidgetWin_1
;			SendInput, {PGDN}
			wb := IEGet("FashionGo Vendor Admin - Internet Explorer")
			wb.document.getElementsByTagName("TEXTAREA")[1].focus()
			Send, {Down}
			
			
			; 여러 메모들 띄우기
			SoundBeep, 750, 500
			MsgBox, 4100, Memo, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nThe Shipping State is %State%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%

			; No 눌렀으면 다음 주문으로 이동
			IfMsgBox, No
			{


				
				; 펜딩 오더가 있어서 Create Invoice 탭을 열고 Account Summary 창을 열어봤을 경우엔 일단 그 창들 닫고 처리해야 되니까
				IfWinExist, Customer Order +Zoom In
					WinClose
				
				IfWinExist, Accounts Summary
					WinClose


				; PO_Number 나 Order_ID_Only 값이 있으면 어차피 다음 Order Id 번호로 넘어가는 게 아니니 어플 다시 시작하기
				if(PO_Number || Order_ID_Only)
					Reload

				
				WinClose, ahk_exe iexplore.exe
				;~ WinClose, ahk_class Chrome_WidgetWin_1
				CurrentOrderIdNumber++
				continue
			}
			
			; Yes 눌렀으면 인보이스 뽑기
			IfMsgBox, Yes
			{

				; 펜딩 오더가 있어서 Create Invoice 탭을 열고 Account Summary 창을 열어봤을 경우엔 일단 그 창들 닫고 처리해야 되니까
				IfWinExist, Customer Order +Zoom In
					WinClose
				
				IfWinExist, Accounts Summary
					WinClose
				
				; BO 목록표(인보이스 종이) 인쇄하기
				PrintingBOList()			

				; 목록표 인쇄 후 Create Invoce 탭으로 넘어가서 인보이스 만들 준비
				OpenCreateInvoiceTab()
				Sleep 200
								
				;New & Clear 버튼 클릭
				MouseClick, l, 60, 124, 2, 
				sleep 1000			
											
				; CompanyName 변수값 넣기
				DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
				Sleep 100
				MouseGetPos, , , , control, 1
				ControlSetText, %control%, %CompanyName%, %wintitle%
				ControlClick %control%, %wintitle%
				SendInput, {Enter}
				

				; Warning 창이 뜨면 Credit 있다는 것
				Sleep 2000
				IfWinActive, Warning
				{					
					MsgBox, IT'S APPLY CREDIT
					WinClose, Warning
				}				

				; Sales Orders 버튼 클릭 하기
				MouseClick, l, 232, 388
				Sleep 100
				
				
				; Please enter Company Nmae first 에러 창이 뜨면 창 닫고 CompanyName 다시 입력하기
				Sleep 500
				ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[ERROR]COMPANY_NAME_FIRST.png
				
				if(ErrorLevel = 0){			
					ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[ERROR]COMPANY_NAME_FIRST_OK.png
					MouseClick, l, %FoundX%, %FoundY%
					Sleep 200			

					; CompanyName 변수값 넣기
					DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
					Sleep 100
					MouseGetPos, , , , control, 1
					ControlSetText, %control%, %CompanyName%, %wintitle%
					ControlClick %control%, %wintitle%
					SendInput, {Enter}
				}


				; The data no found 에러 떴을 때. 아마 같은 회사명 가진 다른 회사가 있을 것임
				IfWinExist, Confirm (No Data)
					MsgBox, 262144, Having Same Name, IT'S AN ERROR OCCURED, PROBABLY THERE IS AN ANOTHER COMPANY HAVING SAME COMPANY NAME.`n`nPLEASE CHECK IT FIRST AND THEN CLICK OK BUTTON TO CONTINUE
				
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


				; 다음 주문으로 넘어가기 전 메세지 띄우기
				MsgBox, 262144, Memo, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nThe Shipping State is %State%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%

				
				; PO_Number 나 Order_ID_Only 값이 있으면 어차피 다음 Order Id 번호로 넘어가는 게 아니니 어플 다시 시작하기
				if(PO_Number || Order_ID_Only)
					Reload

				
				WinClose, Transfer from Sales Order
				WinClose, Accounts Summary
				
				
				WinClose, ahk_exe iexplore.exe
				;~ WinClose, ahk_class Chrome_WidgetWin_1
				CurrentOrderIdNumber++
				continue	

				
			}
			
			
			; PO_Number 나 Order_ID_Only 값이 있으면 어차피 다음 Order Id 번호로 넘어가는 게 아니니 어플 다시 시작하기
			if(PO_Number || Order_ID_Only)
				Reload
						
			
		}



		;#####################################
		;펜딩 값이 없었을 경우
		;현재 Account Summary 창이 열려있는 상태
		;#####################################

		; BO 목록 인쇄할지 말지 묻고 처리하기
		; 인쇄할 지 말지 판단하기 위해 일단 BO 목록표 열기 (Customer Order + 버튼 클릭)
		ControlClick, WindowsForms10.BUTTON.app.0.378734a4, Accounts Summary
		WinWaitActive, Customer Order +Zoom In
		WinMaximize
		


		; QOH 더블클릭해서 오름차순으로 정열하기
		Text:="|<QOH>*161$24.T3sVlaAVVYAVUY4VUY4zUY4VVYAVlaAVT3sV20003U00U"
		if ok:=FindText(639,99,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
			Click
		}


		; Order Date 위치 찾은 후 Capture2Text 이용해서 주문 날짜 값 얻기
		; 클립보드에서 변수 값을 읽는 시간이 길어서 
		; function.ahk 의 FindingOrderMonthAndCompareWithNow() 함수에 관련 내용을 집어넣었음
		
/*		
		; QOH 더블클릭해서 오름차순으로 정열하기
		; Color 버튼의 117만큼 오른쪽으로 옆을 클릭하기
		ImageSearch, FoundX, FoundY, 416, 60, 626, 176, %A_ScripDir%PICTURES\ColorButtonONCustomerOrderZoonIn.png	
		FoundX += 117
		FoundY += 5
		MouseClick, l, %FoundX%, %FoundY%, 2
*/


		;~ ##################################################################################################################################################################################################################
		;~ #################################################################    아이템 찾는 이건 나중에 해보자    ##############################################################################################################
		; 열려있는 Customer Order + 창에서 정보 읽기
		;~ ReadingINFOofCustomerOrderWindow()
		;~ ##################################################################################################################################################################################################################		
		;~ ##################################################################################################################################################################################################################		

		

		; 패션고 페이지 활성화 시켜서 스크롤 다운 한 번 해주기
		; 이러면 주문 어떻게 했는지 보고 판단할 때 도움되니까
;		WinActivate, ahk_exe iexplore.exe
		;~ WinActivate, ahk_class Chrome_WidgetWin_1
;		SendInput, {PGDN}
		wb := IEGet("FashionGo Vendor Admin - Internet Explorer")
		wb.document.getElementsByTagName("TEXTAREA")[1].focus()
		Send, {Down}

		
		; BO 목록 인쇄할 지 묻기. Yes 누르면 인쇄
		SoundBeep, 750, 500
		MsgBox, 4100, , The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nThe Shipping State is %State%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%
		; No 눌렀으면 다음 주문으로 이동
		IfMsgBox, No
		{
			
			; PO_Number 나 Order_ID_Only 값이 있으면 어차피 다음 Order Id 번호로 넘어가는 게 아니니 어플 다시 시작하기
			if(PO_Number || Order_ID_Only)
				Reload

			
			Sleep 100
			WinClose, Customer Order +Zoom In
			WinClose, Accounts Summary
			
			WinClose, ahk_exe iexplore.exe
			;~ WinClose, ahk_class Chrome_WidgetWin_1
			CurrentOrderIdNumber++
			continue
		}
		
		; BO 목록표 인쇄하기
		; PrintingBOList 함수 다른데서도 사용해야 되니까 일단 Customer Order +Zoom In 창과 Account Summary 창 닫고 다시 처음부터 인쇄 프로세스 시작해서 BO 목록(인보이스 종이) 인쇄하기
		WinClose, Customer Order +Zoom In
		WinClose, Accounts Summary
		PrintingBOList()
		
		; 목록표 인쇄 후 Create Invoce 탭으로 넘어가서 인보이스 만들 준비
		OpenCreateInvoiceTab()
		Sleep 200
						
		;New & Clear 버튼 클릭
		MouseClick, l, 60, 124, 2, 
		sleep 1000			
									
		; CompanyName 변수값 넣기
		DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlSetText, %control%, %CompanyName%, %wintitle%
		ControlClick %control%, %wintitle%
		SendInput, {Enter}
		

		; Warning 창이 뜨면 Credit 있다는 것
		Sleep 1500
		IfWinActive, Warning
		{					
			MsgBox, IT'S APPLY CREDIT
			WinClose, Warning
		}

		; Sales Orders 버튼 클릭 하기
		MouseClick, l, 232, 388
		Sleep 100
	
		
		; Please enter Company Nmae first 에러 창이 뜨면 창 닫고 CompanyName 다시 입력하기
		Sleep 500
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[ERROR]COMPANY_NAME_FIRST.png
		
		if(ErrorLevel = 0){			
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %A_ScripDir%PICTURES\[ERROR]COMPANY_NAME_FIRST_OK.png
			MouseClick, l, %FoundX%, %FoundY%
			Sleep 200			

			; CompanyName 변수값 넣기
			DllCall("SetCursorPos", int, 106-8, int, 301-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
			Sleep 100
			MouseGetPos, , , , control, 1
			ControlSetText, %control%, %CompanyName%, %wintitle%
			ControlClick %control%, %wintitle%
			SendInput, {Enter}
		}
		
	
		; The data no found 에러 떴을 때. 아마 같은 회사명 가진 다른 회사가 있을 것임
		IfWinExist, Confirm (No Data)
			MsgBox, 262144, Having Same Name, IT'S AN ERROR OCCURED, PROBABLY THERE IS ANOTHER COMPANY HAVING SAME COMPANY NAME.`n`nPLEASE CHECK IT FIRST AND THEN CLICK OK BUTTON TO CONTINUE

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


		; 다음 주문으로 넘어가기 전 메세지 띄우기
		MsgBox, 262144, Memo, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nREADY TO MOVE TO NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%
		

		; PO_Number 나 Order_ID_Only 값이 있으면 어차피 다음 Order Id 번호로 넘어가는 게 아니니 어플 다시 시작하기
		if(PO_Number || Order_ID_Only)
			Reload
		
		
		; 혹시 창을 열어놓은 상태로 OK 버튼 눌렀을 수도 있으니까 닫기
		WinClose, Transfer from Sales Order
		
		WinClose, ahk_exe iexplore.exe
		;~ WinClose, ahk_class Chrome_WidgetWin_1
		CurrentOrderIdNumber++
		continue
		
		
	}



	; gui에서 PO 번호 입력했으면 그 번호로 Order ID 검색 후 열기
	ProcessWhenPoNumberInputted(PO_Number){
		
		;~ OpenCreateSalesOrdersSmallTab()
		
		Send, {F3}
		WinWaitActive, Find

		; 커서 상태가 작업처리중이면 끝날때까지 기다리기
		while (A_cursor = "Wait")
			Sleep 500

		Sleep, 500
		
		ControlSetText, WindowsForms10.EDIT.app.0.378734a4, Cust PO NO, Find
		;~ ControlSetText, WindowsForms10.EDIT.app.0.378734a4, {Enter}, Find
		
		ControlClick, WindowsForms10.Window.8.app.0.378734a7, Find, ,l
		Send, {Down}
		Send, {Down}
		Send, {Enter}
		
		
		Send, {Tab}
		Send, %PO_Number%
		Send, {Enter}
		Send, {Enter}



		Sleep 2000 ; 위의 동작으로 검색한 뒤에는 이정도 시간을 둬야됨





		Text:="|<CUST PO NO>*126$63.D0083kQ0l3W8010F4E68WUF7S2910t8A2990F88591UF882910Z8A28t0S884d1UF182110X8++M90E4E4MWDDCC20Q0V3Y"

		if ok:=FindText(697,417,150000,150000,0,0,Text)
		{
			;~ MsgBox, found it
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5

	
			; PO Number 입력해서 Order ID 찾기 위한 위치로 이동해서 더블 클릭하기
			MouseClick, left, X, Y+20, 2
			
			; PO Number 입력
;			Send, % PO_Number
			
			Sleep, 700
			
			; 찾은 결과 더블 클릭해서 페이지 열기
			MouseMove, X, Y+42
			; 현재 마우스 위치에 더블클릭 합니다 (아래 세 줄 모두 현재 위치에서 더블 클릭하는 것 안 먹어서 그냥 다 해봄)
			Sleep 200
			MouseClick, left
			MouseClick, left, , , 2
			Click 2



			; 3초 동안 Find 창 닫히기 기다리기
			WinWaitClose, Find, , 3
			
			; Find 안 닫혔으면, 즉 더블 클릭해서 주문 페이지로 넘어가지 않았으면 Find 창 닫고 처음부터 다시 시작하기
			IfWinExist, Find
			{
				WinClose, Find
				ProcessWhenPoNumberInputted(PO_Number)
			}
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
			while (A_cursor = "Wait")
				Sleep 500

			Sleep, 500
		}		

		return
	}
	
	
Exitapp	


GuiClose:
 ExitApp

Esc::
 Exitapp

^e::
CompanyName = %Clipboard%
StringUpper, CompanyName, CompanyName ; Staff only notes 대문자로 바꾸기
SendInput, %CompanyName%
return

^w::
;~ ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, %A_ScripDir%PICTURES\[FGNewOrder01]TabButtonNextToPONumber.png
;~ FoundX += 5
;~ MouseClick, l, %FoundX%, %FoundY%	
SendInput, %CurrentPONumber%
return

