#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include OpenFGToVoidBO.ahk



global CurrentOrderIdNumber, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, CustomerNoteOnWebVal, StaffOnlyNoteVal, CompanyName, PendingOrderStatus, AlreadyProcessedItem, State, ClickLOC
global Style_NO, Style_Color


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


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
	BlockInput, Mouse


	Array := object()

	RESTART:
	Loop
	{
		
		
		WinClose, Details Edit
		
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


		; 메모가 들어갈 변수값 초기화
		CustomerMemoOnLAMBS :=
		SalesOrderMemoONLAMBS :=
		StaffOnlyNoteVal := 
		ClickLOC :=
		

		; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		IfWinNotExist, ahk_class XLMAIN
		{
			MsgBox, 262144, No Excel file Warning, Please Open an Excel File of BO list
			continue
		}

		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible


		; 만약 열려있는 파일이 쓸 데 없는 값을 갖고 있으면 정리하고 시작하기
		; 파일의 B1 셀에 Jodifl 이 들어있으면 앞의 8줄 지우기
		ValofB1 := Xl.Range("B1").Value
		IfEqual, ValofB1, Jodifl
			Xl.Sheets(1).Range("A1:A8").EntireRow.Delete


		;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
		XL_Handle(XL,1) ;get handle to Excel Application
		i := XL_Last_Row(XL)
;		MsgBox % "last row: " XL_Last_Row(XL)  ;Last row
;		MsgBox, % i

		; 엑셀창 최소화하기
		;~ WinMinimize, ahk_class XLMAIN

		; 엑셀에 값이 들어간 만큼(i 값 만큼) 루프 돌면서 엑셀에서 값 읽기
		Loop, %i%{
		;Loop{
		
						
			/* 배열로부터 읽기 첫 번째 방법
			Loop % Array.Maxindex(){
				MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
			}
			*/

		
			; Order ID 값은 D Column 에 있음
			; 앞에서 쓸 데 없는 값을 지워줬으니 D1 에 지금 사용 할 Order ID 값 있음			
;			RawOrderID := Xl.Range("D1").Value
			RawOrderID := Xl.Range("C1").Value
			Style_NO := Xl.Range("H1").Value
			Style_Color := Xl.Range("I1").Value
			
			;소수점 뒷자리 정리
			RegExMatch(RawOrderID, "imU)(\d*)\.", SubPat)
			
			; 정리된 값 RefinedOrderID 에 넣기
			RefinedOrderID := SubPat1
			
			
			; 에러 확인용
;			MsgBox, RefinedOrderID is : %RefinedOrderID%
					
			
			
			; 만약 지금 얻은 RefinedOrderID 값이 이전 Order ID 값을 저장하고 있는 previousNumber 값과 같다면 
			; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
;			IfEqual, RefinedOrderID, %previousNumber%		
;			{
				;MsgBox, duplicated number
;				Xl.Sheets(1).Range("A1").EntireRow.Delete
;				continue
;			}
			
			
	
			
	;		MsgBox, %RefinedOrderID%
			
			; 첫 번째 Row 값은 변수에 넣었으니 엑셀에서 지워주기
;			Xl.Sheets(1).Range("A1").EntireRow.Delete
			
			
			; 중복되는 값의 비교를 위해 previousNumber 변수에 RefinedOrderID 값 넣기
			previousNumber := RefinedOrderID
			
			
			
			; RefinedOrderID 값을 넘겨주면서 BO_LAMBSProcessing 함수 호출하기
			BO_LAMBSProcessing(RefinedOrderID)




		}

	}








BO_LAMBSProcessing(CurrentOrderIdNumber)
{

; 혹시 값을 못 얻어서 다시 입력해야 되는 경우만을 위한 go to 회돌이
The_Biginning:
	

	Start()

	wintitle = LAMBS -  Garment Manufacturer & Wholesale Software
	
	; 메모리에 있는 값, PendingOrderStatus 값, AlreadyProcessedItem 값 초기화 하기
	Clipboard :=
	PendingOrderStatus :=
	AlreadyProcessedItem :=


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



	
	
	; Order ID 입력칸으로 이동 후 CurrentOrderIdNumber 변수값 넣기
	DllCall("SetCursorPos", int, 181-8, int, 190-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;	MouseMove, 181, 190
	Sleep 100
	MouseGetPos, , , , control, 1
	ControlSetText, %control%, %CurrentOrderIdNumber%, %wintitle%
	ControlClick %control%, %wintitle%
	SendInput, {Enter}


	Sleep 1000


	; Company Name 얻기
	DllCall("SetCursorPos", int, 87-8, int, 375-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;	MouseMove, 87, 375
	Sleep 100
	MouseGetPos, , , , control, 1
	ControlGetText, CompanyName, %control%, %wintitle%

	
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
	State_Abbreviations()

	
	; PO 번호 CurrentPONumber 변수에 저장하기	
	DllCall("SetCursorPos", int, 994-8, int, 378-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
;	MouseMove, 994, 378
	Sleep 100
	MouseGetPos, , , , control, 1
	ControlGetText, CurrentPONumber, %control%, %wintitle%



	; 혹시 어떤 오류 때문에 Order ID를 입력 못해서 PO값이 없으면 다시 처음부터 시작해서 값 입력하기
	if(!CurrentPONumber)
		goto, The_Biginning
	
	
	; 해당 PO 가 패션고 페이지면 페이지 열기
	if CurrentPONumber contains MTR
	{
		MsgBox, 262144, Memo, Click OK to Open FG Page
		OpenFGToVoidBO(CurrentPONumber)
	}
	
/*
	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible


	; 첫 번째 Row 값은 이미 읽어서 사용했으니 지워주기
	Xl.Sheets(1).Range("A1").EntireRow.Delete
*/

	;~ MsgBox, 262144, Memo, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nREADY TO MOVE TO NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%`n`n%AdditionalInfo%
	MsgBox, 262144, Memo, continue
	
	
	
	
}


	
Exitapp



GuiClose:
 ExitApp

Esc::
 Xl.ActiveWorkbook.save()
 Exitapp

^W::
CompanyName = %Clipboard%
StringUpper, CompanyName, CompanyName ; Staff only notes 대문자로 바꾸기
SendInput, %CompanyName%
return


^Q::
SendInput, {BackSpace}
Send, {Space}
SendInput, VOIDED BC OLD BO 12/18/17 CH
return

^Space::
Xl.Sheets(1).Range("A1").EntireRow.Delete
	GroupAdd,ExplorerGroup, ahk_class IEFrame
	WinClose,ahk_group ExplorerGroup
	Sleep 150
	Send, {ENTER}
	winclose, Memo
	Sleep 300
	Reload
return