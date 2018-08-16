#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
/*
Arr_ADD := object()
EntireAddress = 2356 E 7975 S #203, South Weber, UT 84405, United States

;~ word_array := StrSplit(EntireAddress, A_Space, ".")  ; 점은 제외합니다.
Arr_ADD := StrSplit(EntireAddress, ",")  ; 콤마 나올때마다 문자열 나누기


; ADD2 찾아서 배열 5번째에 넣고 1번째에는 ADD1만 남기기
UnquotedOutputVar = im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)
Arr_ADD[5] := FindAdd2_In_Add1(Arr_ADD[1], UnquotedOutputVar) ; Arr_ADD[1] 에 들어있는 전체 주소를 넘겨서 ADD2 만 Arr_ADD[5] 에 넣기
Arr_ADD[1] := DeleteAdd2_In_Add1(Arr_ADD[1], UnquotedOutputVar) ; 전체 주소 중 ADD2를 지운뒤 Arr_ADD[1]에 넣기


; ZIP 찾아서 배열 6번째에 넣고 3번째에는 State(州)만 넣기
UnquotedOutputVar = im)(\d.*)
Arr_ADD[6] := FindAdd2_In_Add1(Arr_ADD[3], UnquotedOutputVar) ; Arr_ADD[1] 에 들어있는 전체 주소를 넘겨서 ADD2 만 Arr_ADD[5] 에 넣기
Arr_ADD[3] := DeleteAdd2_In_Add1(Arr_ADD[3], UnquotedOutputVar) ; 전체 주소 중 ADD2를 지운뒤 Arr_ADD[1]에 넣기

CheckPendingIDExist := RegExReplace(CheckPendingIDExist, "[^0-9]*", "")



Loop % Arr_ADD.Maxindex(){
	Arr_ADD[A_Index] := Trim(Arr_ADD[A_Index])
	MsgBox % "Element number " . A_Index . " is |" . Arr_ADD[A_Index] . "|"
}
MsgBox puase


FindAdd2_In_Add1(Arr_Original, UnquotedOutputVar){	
	;~ while(RegExMatch(Arr_Original, "im)\s((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)", FoundAdd2)){
	while(RegExMatch(Arr_Original, UnquotedOutputVar, FoundAdd2)){
		if(ErrorLevel = 0){
			Temp := FoundAdd2
			Temp := Trim(Temp)
			;~ MsgBox, % Temp . "||"
		}
	break
	}
	return Temp
}


DeleteAdd2_In_Add1(Arr_Original, UnquotedOutputVar){
	;~ while(FoundPos := RegExMatch(Arr_Original, "im)((unit|Suite|Ste|#|Apt|SPACE|BLDG|Building|Sujite).*)", FoundAdd2)){
	while(FoundPos := RegExMatch(Arr_Original, UnquotedOutputVar, FoundAdd2)){
		if(ErrorLevel = 0){
			Temp := SubStr(Arr_Original, 1, FoundPos-1)
			Temp := Trim(Temp)
			;~ MsgBox, % Temp . "||"
		}
	break
	}
	return Temp
}
*/





#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk


global CurrentOrderIdNumber, POSourceOrMemo, CustomerMemoOnLAMBS, SalesOrderMemoONLAMBS, CustomerNoteOnWebVal, StaffOnlyNoteVal, CompanyName, PendingOrderStatus, AlreadyProcessedItem, State, ClickLOC

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





	Loop
	{

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
			
			; Order ID 값은 C Column 에 있음
			; 앞에서 쓸 데 없는 값을 지워줬으니 C1 에 지금 사용 할 Order ID 값 있음
			RawOrderID := Xl.Range("C1").Value
			
			;소수점 뒷자리 정리
			RegExMatch(RawOrderID, "imU)(\d*)\.", SubPat)
			
			; 정리된 값 RefinedOrderID 에 넣기
			RefinedOrderID := SubPat1
			
			; 에러 확인용
;			MsgBox, RefinedOrderID is : %RefinedOrderID%
			
			

			; 만약 RefinedOrderID 에 값이 없으면 파일이 끝났거나 이전에 사용했던 파일이 안 닫히고 열려있는 것이므로 파일 닫고 프로그램 다시 시작하기
			if(!RefinedOrderID)
			{				
				; 파일이 끝났으니 메세지 띄우고 프로그램 다시 시작하기
				MsgBox, 262144, Old File Notification, ALL ORDER ID NUMBERS HAVE BEEN PROCESSED`nPLEASE OPEN NEW BO LIST EXCEL FILE
				
				
				; 저장 않고 종료하는 법을 못 찾아서 그냥 일단 임시로 저장 후 바로 지우기		
				path = %A_ScriptDir%\CreatedFiles\temporary.xls
				XL.ActiveWorkbook.SaveAs(path) ;'path' is a variable with the path and name of the file you desire
				
				; 엑셀 종료하기
				;xL.ActiveWorkbook.SaveAs("testXLfile",56)               ;51 is an xlsx, 56 is an xls
				xl.WorkBooks.Close()                                    ;close file
				xl.quit
				
				; 방금 만든 파일 지우기
				FileDelete, %A_ScriptDir%\CreatedFiles\temporary.xls
				
				; 프로그램 재시작
				Reload
			}
				
			
			
			
			
			; 만약 지금 얻은 RefinedOrderID 값이 이전 Order ID 값을 저장하고 있는 previousNumber 값과 같다면 
			; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
			IfEqual, RefinedOrderID, %previousNumber%
			{
				;MsgBox, duplicated number
				Xl.Sheets(1).Range("A1").EntireRow.Delete
				continue
			}


	;		MsgBox, %RefinedOrderID%
			
			; 첫 번째 Row 값은 변수에 넣었으니 엑셀에서 지워주기
			Xl.Sheets(1).Range("A1").EntireRow.Delete
			
			
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


	
;	MsgBox, 4100, Briff Memo, %CompanyName%`n`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%CurrentPONumber%



	; POSourceOrMemo 값에 TIFFANY 가 들어있으면 TIFFANY ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
	IfInString, POSourceOrMemo, TIFFANY
	{
		;MsgBox, , , IT'S TIFFANY ONLY ORDER, MOVE TO NEXT ORDER`n`n`nTHIS WINDOW WILL BE CLOSED IN 5 SECONDS, 5
		MsgBox, , , IT'S TIFFANY ONLY ORDER, MOVE TO NEXT ORDER, 3
		return
	}


	; POSourceOrMemo 값에 JONATHAN 가 들어있으면 JONATHAN ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
	IfInString, POSourceOrMemo, JONATHAN
	{		
		MsgBox, , , IT'S JONATHAN ONLY ORDER, MOVE TO NEXT ORDER, 3
		return
	}

	; POSourceOrMemo 값에 CONNIE 가 들어있으면 CONNIE ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
	IfInString, POSourceOrMemo, CONNIE
	{		
		MsgBox, , , IT'S CONNIE ONLY ORDER, MOVE TO NEXT ORDER, 3
		return
	}
	
	; POSourceOrMemo 값에 ROONY 가 들어있으면 ROONY ONLY 이므로 처음으로 돌아가 회돌이 다시 시작하기
	IfInString, POSourceOrMemo, ROONY
	{		
		MsgBox, , , IT'S ROONY ONLY ORDER, MOVE TO NEXT ORDER, 3
		return
	}
	
	

	; Account Summary 띄워서 Pending Order 있는지 확인 하기
	AccountSummayrProcessingONCreateSalesOrdersSmallTab()
	
	
	; 만약 펜딩값이 있었으면 AccountSummayrProcessingONCreateSalesOrdersSmallTab 함수 안에서 Create Invoice 창으로 넘어가는 처리 하고
	; PendingOrderStatus 변수에 값을 넣어서 펜딩값이 있어서 처리해줬다는 표시를 해줬음
	
	; 그 후처리 하기


	;####################
	;펜딩 값이 있었을 경우
	;####################
	
	if(PendingOrderStatus){
		
		; 여러 메모들 띄우기
		SoundBeep, 750, 500
		MsgBox, 4100, Memo, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nThe Shipping State is %State%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
		
		IfMsgBox, No
		{
			; 펜딩 오더가 있어서 Create Invoice 탭을 열고 Account Summary 창을 열어봤을 경우엔 일단 그 창들 닫고 처리해야 되니까
			IfWinExist, Customer Order +Zoom In
				WinClose
			
			IfWinExist, Accounts Summary
				WinClose

			return
		}
		
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
			MsgBox, 262144, Move to the next, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nThe Shipping State is %State%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
			
			WinClose, Transfer from Sales Order
			
			
			return

			
		}
	}


	;#####################################
	;펜딩 값이 없었을 경우
	;현재 Account Summary 창이 열려있는 상태
	;#####################################



;~ ################## 아이템별로 만들땐 이게 유용한데 예를들어 날짜별로 하면 없는 아이템도 있을 수 있기 때문에 이것이 적용이 안됨 ###############################
/*

	; 만약 모든 메모 변수에 값이 없다면(즉, 특별한 주문이 없다면) 그냥 인보이스 인쇄하기
	; 왜냐면 백오더는 일단 물건이 들어온 상태이기 때문에 기본적으로 물건을 빼기 때문
	if(!CustomerMemoOnLAMBS & !SalesOrderMemoONLAMBS & !StaffOnlyNoteVal){
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
		
		; Sales Orders 버튼 클릭 하기
		MouseClick, l, 232, 388


		; 다음 주문으로 넘어가기 전 메세지 띄우기
		MsgBox, 262144, Move to the next, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nREADY TO MOVE TO NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
		
		; 혹시 창을 열어놓은 상태로 OK 버튼 눌렀을 수도 있으니까 닫기
		WinClose, Transfer from Sales Order

		return
	}
*/	


	; BO 목록 인쇄할지 말지 묻고 처리하기	
	; 인쇄할 지 말지 판단하기 위해 일단 BO 목록표 열기 (Customer Order + 버튼 클릭)
	ControlClick, WindowsForms10.BUTTON.app.0.378734a4, Accounts Summary
	WinWaitActive, Customer Order +Zoom In
	WinMaximize
	
/*
	; QOH 더블클릭해서 오름차순으로 정열하기
	ImageSearch, FoundX, FoundY, 416, 60, 626, 176, %A_ScripDir%PICTURES\ColorButtonONCustomerOrderZoonIn.png	
	FoundX += 117
	FoundY += 5
	MouseClick, l, %FoundX%, %FoundY%, 2
*/

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
		
	; BO 목록 인쇄할 지 묻기. Yes 누르면 인쇄
	SoundBeep, 750, 500
	MsgBox, 4100, , The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nThe Shipping State is %State%`nWOULD YOU LIKE TO PRINT OUT BO LIST?`nIF YOU CLICK No, MOVE TO THE NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
	; No 눌렀으면 다음 주문으로 이동
	IfMsgBox, No
	{
		Sleep 100
		WinClose, Customer Order +Zoom In
		WinClose, Accounts Summary
		
		return
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
	MsgBox, 262144, Move to the next, The Number is %CurrentOrderIdNumber%[%ClickLOC%]`nREADY TO MOVE TO NEXT ORDER`n`n%POSourceOrMemo%`n`n`n%CustomerMemoOnLAMBS%`n`n`n%SalesOrderMemoONLAMBS%`n`n`n%StaffOnlyNoteVal%
	
	; 혹시 창을 열어놓은 상태로 OK 버튼 눌렀을 수도 있으니까 닫기
	WinClose, Transfer from Sales Order
	
		
	return
}


	
	
	
Exitapp



GuiClose:
 ExitApp

Esc::
 Xl.ActiveWorkbook.save()
 Exitapp

^e::
CompanyName = %Clipboard%
StringUpper, CompanyName, CompanyName ; Staff only notes 대문자로 바꾸기
SendInput, %CompanyName%
return


^w::
Send, %RefinedOrderID%
return