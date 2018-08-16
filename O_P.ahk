; ######################################################################
; 지금 오픈되어 있는 오더들 엑셀에서 읽어와서 오래된 펜딩오더 찾기
/*
( open_pick_qty > 0 )  
*/
; ######################################################################



#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\


#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk


;~ #Include N41.ahk
#Include CommN41.ahk



global #ThatCurrentlyUsing




	MsgBox, 262144, REFINE ALLOCATION FILE, CLICK OK TO CONTINUE


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative


	; 사용자의 마우스 이동 막음
	BlockInput, MouseMove	






	;~ N_driver := new N41
	N_driver := new CommN41


	#ThatCurrentlyUsing = 2
	Loop{
		
		; 펜딩 오더 엑셀에서 고객 코드 읽어오기
		custCode := getCustomerCodeFromExcelThatTheListOfPendingOrders(#ThatCurrentlyUsing)
;		MsgBox, % "||" . custCode . "||"
		
		
		
			WinActivate, ahk_class FNWND3126
			Sleep 500
			
			; 왼쪽 메뉴바에 있는 Customer 클릭하기
			N_driver.ClickCustomerMarkOnTheLeftBar()			
			
			
			; 왼쪽 메뉴바에 있는 SO Manager 클릭하기
			N_driver.ClickSOManagerOnTheLeftBar()			
			
			
			; SO Manager 에 있는 Customer 표시 찾기. 고객 코드 입력하기 위함
			N_driver.FindCustomerMarkToFillInTheBlank()
			
			
			; 검색창에 고객 코드 입력 후 엔터쳐서 검색하기		
			Sleep 300
			Send, ^a
			Sleep 150
			Send, % CustCode
			Sleep 150
			SendInput, {Enter}
			Sleep 150
			
			
			; 커서 상태가 작업처리중이면 끝날때까지 기다리기
/*			
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}
			Sleep 300
			Loop{
				if(A_cursor = "Wait"){
					Sleep 500
				}
				else
					break
			}			
*/			
			
			while (A_cursor = "Wait")
				Sleep 1000
			Sleep 1000
			
			while (A_cursor = "Wait")
				Sleep 1000
			Sleep 1000
			
			while (A_cursor = "Wait")
				Sleep 1000
			Sleep 1000
			
			while (A_cursor = "Wait")
				Sleep 1000
			Sleep 1000
			
			while (A_cursor = "Wait")
				Sleep 1000
			Sleep 1000

		
		
			; 펜딩 오더가 있는지 확인키 위한 변수.
			isTherePendingOrder = 0		
		
		
		
			; pick ticket 섹션에서 pick date 읽어서 변수에 저장
			pickDate := N_driver.getPickDateOnPickTicketSectionOfSOManager()
			
;			MsgBox, % pickDate


			; 알파벳과 숫자만 저장 (알페벳과 숫자 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
			pickDate := RegExReplace(pickDate, "[^a-zA-Z0-9]", "")


			; 오늘 날짜
			todaysDate = %A_MM%%A_DD%%A_YYYY%  ; ############ 실전에는 이걸 써야됨 ################
			;~ todaysDate = 05212018


			; pickDate 변수에서 뒤에 있는 시간(00:00:00 이렇게 표시됨)을 제외한 년월일만 뽑아서 변수에 다시 넣기
			StringLeft, pickDate, pickDate, 8


			;~ MsgBox, % "마지막 배송 날짜 : "pickDate . "`n" . "            오늘 날짜 : " . todaysDate



			; pick date 가 없으면, 즉 현재 열려있는 오더가 없으면 
			IF(!pickDate)
			{
				doesItHaveToBeDeleted = 1
;				MsgBox, 262144, Title, 배송 날짜 없음 목록에서 지워야 됨 `n`n%pickDate%
				deleteCustInfoFromExcel(#ThatCurrentlyUsing) ; 엑셀에서 고객 코드가 있는 해당 row 지우기
				;~ #ThatCurrentlyUsing++ ; 다음 코드로 넘어가기 위해 1 증가
				continue ; loop 맨 위로 다시 올라가기 위해 
			}



		; 펜딩 날짜가 있을때
		IF(pickDate)
		{
			;~ MsgBox, 262144, Title, 펜딩 날짜 있음
			
			; 연도를 변수에 저장
			yearOfpickDate := SubStr(pickDate, 5, 4)
			yearOfToday := SubStr(todaysDate, 5, 4)

			dateOfpickDate := SubStr(pickDate, 3, 2)
			dateOfToday := SubStr(todaysDate, 3, 2)

			monthOfpickDate := SubStr(pickDate, 1, 2)
			monthOfToday := SubStr(todaysDate, 1, 2)
						

;	MsgBox, % "yearOfpickDate : " . yearOfpickDate . "`n" . "yearOfToday : " . yearOfToday . "`n`n" . "dateOfpickDate : " . dateOfpickDate . "`n" . "dateOfToday : " . dateOfToday . "`n`n" . "monthOfpickDate : " . monthOfpickDate . "`n" . "monthOfToday : " . monthOfToday


				if(yearOfpickDate != yearOfToday){

					doesItHaveToBeDeleted = 0
;					MsgBox, 262144, Title, 연도가 같지 않음. 확인하기 위해 목록에 남겨놔야 됨
					#ThatCurrentlyUsing++ ; 다음 코드로 넘어가기 위해 1 증가
					continue ; loop 맨 위로 다시 올라가기 위해 
				}

				; 펜딩 날짜의 달과 오늘 날짜의 달이 같지 않으면 옛날 주문이 확실하니 남겨놔야 됨
				if(monthOfpickDate != monthOfToday){
					
					doesItHaveToBeDeleted = 0
;					MsgBox, 262144, Title, 오늘 날짜와 달이 같지 않음. 확인하기 위해 목록에 남겨놔야 됨
					#ThatCurrentlyUsing++ ; 다음 코드로 넘어가기 위해 1 증가
					continue ; loop 맨 위로 다시 올라가기 위해 
					
				}
				
			

				if(dateOfpickDate >= dateOfToday - 7){

					doesItHaveToBeDeleted = 1
;					MsgBox, 262144, Title, 7일 이내의 최근 날짜. 오래되지 않았으므로 목록에서 지워야 됨
					deleteCustInfoFromExcel(#ThatCurrentlyUsing) ; 엑셀에서 고객 코드가 있는 해당 row 지우기
					;~ #ThatCurrentlyUsing++ ; 다음 코드로 넘어가기 위해 1 증가
					continue ; loop 맨 위로 다시 올라가기 위해 
				}


				if(dateOfpickDate <= dateOfToday - 7){
					
					doesItHaveToBeDeleted = 0
;					MsgBox, 262144, Title, 7일 보다 오래된 날짜. 확인하기 위해 목록에 남겨놔야 됨
					#ThatCurrentlyUsing++ ; 다음 코드로 넘어가기 위해 1 증가
					continue ; loop 맨 위로 다시 올라가기 위해 
				}




			


		}
		
/*		
		; 목록에서 지울 때
		if(doesItHaveToBeDeleted){
			

			; 열려있는 엑셀 창 사용하기
			Xl := ComObjActive("Excel.Application")
			Xl.Visible := True ;by default excel sheets are invisible		
			;~ GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가
			
			
			custCode := Xl.Range("A2").Value
			
			if(!custCode){
				MsgBox, 262144, END LIST, End of the List 
				Exitapp
			}
						
			
			cell#OfCustCode := "A" . #ThatCurrentlyUsing
			Xl.Sheets(1).Range(cell#OfCustCode).EntireRow.Delete
		}
		else
			#ThatCurrentlyUsing++
*/		
		
		
		; ###################################################################################################################
		; 이거 만들어야 됨
/*		
		; 마지막에 도착하면 루프 끝내고 나가기 
		if(#ofLineOfStyle#AndColor > lastLine#OfRow)
			break		
*/						
		; ###################################################################################################################
		
	}
	
	
	
	
	

	Exitapp

	Esc::
	Exitapp	
		
	
	
	
	
	deleteCustInfoFromExcel(#ThatCurrentlyUsing){
		

			; 열려있는 엑셀 창 사용하기
			Xl := ComObjActive("Excel.Application")
			Xl.Visible := True ;by default excel sheets are invisible		
			;~ GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가
			
			
			custCode := Xl.Range("A2").Value
/*			
			if(!custCode){
				MsgBox, 262144, END LIST, End of the List 
				Exitapp
			}
*/						
			
			cell#OfCustCode := "A" . #ThatCurrentlyUsing
			Xl.Sheets(1).Range(cell#OfCustCode).EntireRow.Delete		
			
;			MsgBox, 지웠음
		
		
		
		
	}
	
	
	
	
	
	; 펜딩 오더 엑셀 파일에서 고객 코드 가져오기
	getCustomerCodeFromExcelThatTheListOfPendingOrders(#ThatCurrentlyUsing){
		

		; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		IfWinNotExist, ahk_class XLMAIN
		{
			loop{			
				
				BlockInput, MouseMoveOff ; 사용자의 마우스 이동 허용
				
				MsgBox, 262144, No Excel file Warning, PLEASE OPEN AN ORDER EXCEL FILE
				
				IfWinExist, ahk_class XLMAIN
				{
					BlockInput, MouseMove  ; 사용자의 마우스 이동 막음
					break
				}
			}
		}




		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible		
		;~ GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가
		
		
		; 고객 정보 읽어올 셀 번호
		cell#OfCustCode := "A" . #ThatCurrentlyUsing
		
		; 고객 정보 읽어서 변수에 저장
		custCode := Xl.Range(cell#OfCustCode).Value
		
		if(!custCode){
			Xl.ActiveWorkbook.save()
			
			BlockInput, MouseMoveOff ; 사용자의 마우스 이동 허용
			
			MsgBox, 262144, END LIST, IT'S SAVED, REACHED TO THE END OF THE LIST
			
			Exitapp
		}
		
		; 고객 정보 리턴
		return custCode
	}
