#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk
#Include [Excel]_InsertORDeleteColumns.ahk

#Include CommEXCEL.ahk



#Include FG.ahk
#Include LAS.ahk

#Include N41.ahk
#Include CommN41.ahk

#Include CNewBrow.AHK ; 열린 크롬 창을 이용하는 것이 아닌 새로 크롬 창을 열어서 컨트롤 하는 메소드 들어있는 파일

#Include ChromeGet.ahk


global infoOnPickTicket_Array



; N41P 를 위한 N41 에서 버튼 누르는 등 처리하는 클래스
; 주로 인쇄하기 동작
class N41_ProcessingForPT extends CommN41{



	; Allocation 에서 pre authorized 버튼 누르고 인쇄하기
	Alloc_Print_WITH_PreAuthorized(CustomerPO){

		
		; Open Allocation 의 Chk 전체 선택하게 하기
		CommN41.Click_Chk_On_OpenAllocation()


		; Che 가 체크 안됐으면 함수 재귀호출해서 다시 시작하기
		IsTheCheChecked = 0
		IsTheCheChecked := CommN41.Che_is_Checked() ; 체크박스가 체크됐으면 1을 리턴
		if(IsTheCheChecked == 1){
			;~ MsgBox, the che 체크됐음 - Alloc_Print_WITH_PreAuthorized(CustomerPO)
		}
		else
			N41_ProcessingForPT.Alloc_Print_WITH_PreAuthorized(CustomerPO)



		; Allocation 오더에서 온 것일때는 왼쪽 밑의 Open Allocation 에 있는 버튼 누르기
		CommN41.ClickCreatePickTicketButtonOnOpenAllocation()
		
/*		
		; Pick Ticket 버튼 누르자마자 나오는 SO Manager 창은 아이템들이 선택이 안됐다는 경고창이므로
		; 안내 메세지 띄우고 재귀호출로 다시 시작
		IfWinActive, SO Manager
		{
			Sleep 300
			Send, {Enter}
			MsgBox, 262144, NO CHECKED ERROR, PLEASE SELECT ITEMS TO MAKE A PICK TICKET
			N41_ProcessingForPT.Alloc_Print_WITH_PreAuthorized(CustomerPO)
		}
*/		
		
	
		; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
		Sleep 4000
		
		; Merge 확인 창
		IfWinActive, SO Manager
		{
			
				
			Sleep 300
			Send, {Left}
			Sleep 500
			Send, {Left}
			Sleep 500
			Send, {Enter}
			Sleep 500
			


			WinWaitActive, SO Manager
			IfWinActive, SO Manager
			{
					
				Sleep 1000
				Send, {Enter}
				Sleep 1000
				
				WinWaitActive, Pick Ticket ; Allocation 경고창
				IfWinActive, Pick Ticket
				{
					Sleep 1000
					Send, {Left}
					Sleep 500
					Send, {Enter}
					Sleep 1000
					
					
					WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
					IfWinActive, Pick Ticket
					{
						Sleep 1000
						Send, {Enter}
						Sleep 1000

						; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지
						IsItFromNewOrder = 0
						IsItFromAllocation = 1
						N41_ProcessingForPT.FromClickingPreAuthorizedButton_To_PrintOutPickTicket(CustomerPO, IsItFromNewOrder, IsItFromAllocation)
						
;MsgBox, 262144, Title, FromClickingPreAuthorizedButton_To_PrintOutPickTicket 함수 끝내고 돌아왔음 MethodsForPI.ahk 파일의 115 번째 줄

; 열려있는 엑셀 창 사용하기
Xl := ComObjActive("Excel.Application")
Xl.Visible := True ;by default excel sheets are invisible

; 첫 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
Xl.Sheets(1).Range("A1").EntireRow.Delete

BO.BOProcess_by_Style(IsItFromAllocation)
						

						
						return
					}					
				}				
			}		
		}
				
		
		WinWaitActive, Pick Ticket ; 확인창
		Sleep 200
		WinActivate, Pick Ticket
		;~ IfWinActive, Pick Ticket
		{
			Sleep 500
			Send, {Enter}
			Sleep 800
				
			; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지
			IsItFromNewOrder = 0
			IsItFromAllocation = 1
			N41_ProcessingForPT.FromClickingPreAuthorizedButton_To_PrintOutPickTicket(CustomerPO, IsItFromNewOrder, IsItFromAllocation)
				
;MsgBox, 262144, Title, FromClickingPreAuthorizedButton_To_PrintOutPickTicket 함수 끝내고 돌아왔음 MethodsForPI.ahk 파일의 137 번째 줄

; 열려있는 엑셀 창 사용하기
Xl := ComObjActive("Excel.Application")
Xl.Visible := True ;by default excel sheets are invisible

; 첫 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
Xl.Sheets(1).Range("A1").EntireRow.Delete

BO.BOProcess_by_Style(IsItFromAllocation)
				
			return
		}


		return
	} ; Alloc_PrintWithPreAuthorized 메소드 끝







	; Allocation 에서 pre authorized 버튼 안 누르고 인쇄하기
	Alloc_Print_WITHOUT_PreAuthorized(CustomerPO){
		
		
		; Open Allocation 의 Chk 전체 선택하게 하기
		CommN41.Click_Chk_On_OpenAllocation()
		

		; Che 가 체크 안됐으면 함수 재귀호출해서 다시 시작하기
		IsTheCheChecked = 0
		IsTheCheChecked := CommN41.Che_is_Checked() ; 체크박스가 체크됐으면 1을 리턴
		if(IsTheCheChecked == 1){
			;~ MsgBox, the che is checked
		}
		else
			N41_ProcessingForPT.Alloc_Print_WITHOUT_PreAuthorized(CustomerPO)

		
		


		; Allocation 오더에서 온 것일때는 왼쪽 밑의 Open Allocation 에 있는 버튼 누르기
		CommN41.ClickCreatePickTicketButtonOnOpenAllocation()		
		
/*		
		; Pick Ticket 버튼 누르자마자 나오는 SO Manager 창은 아이템들이 선택이 안됐다는 경고창이므로
		; 안내 메세지 띄우고 재귀호출로 다시 시작
		IfWinActive, SO Manager
		{
			Sleep 300
			Send, {Enter}
			MsgBox, 262144, NO CHECKED ERROR, PLEASE SELECT ITEMS TO MAKE A PICK TICKET
			N41_ProcessingForPT.Alloc_Print_WITH_PreAuthorized(CustomerPO)
		}
*/				
		
		
	
		; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
		Sleep 4000
		
		; Merge 확인 창
		IfWinActive, SO Manager
		{
			Sleep 300
			Send, {Left}
			Sleep 300
			Send, {Left}
			Sleep 300
			Send, {Enter}
			Sleep 700
			
			WinWaitActive, SO Manager
			IfWinActive, SO Manager
			{
				Sleep 1000
				Send, {Enter}
				Sleep 1000			
				
				WinWaitActive, Pick Ticket ; Allocation 경고창				
				IfWinActive, Pick Ticket
				{
					Sleep 1000
					Send, {Left}
					Sleep 500
					Send, {Enter}
					Sleep 1000
					
					
					WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
					IfWinActive, Pick Ticket
					{
						Sleep 1000
						Send, {Enter}
						Sleep 1000

									
						; pre authorized 버튼 클릭 않고 인쇄하기
						IsItFromExcelFile = 1
						IsItFromAllocation = 1
						N41_ProcessingForPT.PrintWITHOUTClickPreAuthorizedButton(CustomerPO, IsItFromExcelFile, IsItFromNewOrder, IsItFromAllocation)

; 열려있는 엑셀 창 사용하기
Xl := ComObjActive("Excel.Application")
Xl.Visible := True ;by default excel sheets are invisible

; 첫 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
Xl.Sheets(1).Range("A1").EntireRow.Delete

BO.BOProcess_by_Style(IsItFromAllocation)
			
						return
					}				
				}			
			}
		}
			
		
		
		
			
		
		WinWaitActive, Pick Ticket ; 확인창
		Sleep 200
		WinActivate, Pick Ticket
		;~ IfWinActive, Pick Ticket
		{
			Sleep 500
			Send, {Enter}
			Sleep 500			

			; pre authorized 버튼 클릭 않고 인쇄하기
			IsItFromExcelFile = 1
			IsItFromAllocation = 1
			N41_ProcessingForPT.PrintWITHOUTClickPreAuthorizedButton(CustomerPO, IsItFromExcelFile, IsItFromNewOrder, IsItFromAllocation)

; 열려있는 엑셀 창 사용하기
Xl := ComObjActive("Excel.Application")
Xl.Visible := True ;by default excel sheets are invisible

; 첫 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
Xl.Sheets(1).Range("A1").EntireRow.Delete

BO.BOProcess_by_Style(IsItFromAllocation)
						
			return

		}
	
		
		
		return
	} ; Alloc_Print_WITHOUT_PreAuthorized() 메소드 끝









	; 오른쪽 위의 Open SO 화면의 픽티켓 버튼 클릭해서 인쇄하기
	; N41 의 pre authorized 버튼 클릭하기	
	
	; 뉴오더에서 인쇄할 때 Order Status를 New Order 로 바꾸기 위해 표시하기 위한 IsItFromNewOrder
	OpenSO_Print_WITH_PreAuthorized(CustomerPO, IsItFromNewOrder)
	{
		
		; 1사분면의 Pick Ticket 버튼 누르기
		CommN41.ClickCreatePickTicketButton()
		
		
		; 눈에 거슬리지 않게 Gui 최소화 하기
		WinMinimize, Pick Ticket Processing
		WinMinimize, N41 Processing
		
		
		; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
		Sleep 4000
		
		; Merge 확인 창
		IfWinActive, SO Manager
		{
			Sleep 300
			Send, {Left}
			Sleep 500
			Send, {Left}
			Sleep 500
			Send, {Enter}
			Sleep 700
			
			WinWaitActive, SO Manager
			IfWinActive, SO Manager
			{
				Sleep 1000
				Send, {Enter}
				Sleep 1000			
				
				WinWaitActive, Pick Ticket ; Allocation 경고창
				IfWinActive, Pick Ticket
				{
					Sleep 1000
					Send, {Left}
					Sleep 500
					Send, {Enter}
					Sleep 700
					
					
					WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
					IfWinActive, Pick Ticket
					{
						Sleep 1000
						Send, {Enter}
						Sleep 800						
						
						; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지
						N41_ProcessingForPT.FromClickingPreAuthorizedButton_To_PrintOutPickTicket(CustomerPO, IsItFromNewOrder, IsItFromAllocation)
						
						return
					}				
				}			
			}
		}




		WinWaitActive, Pick Ticket ; Allocation 경고창
		;~ IfWinActive, Pick Ticket
		WinActivate, Pick Ticket
		{
;~ MsgBox, Pick Ticket 열렸음			
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
				
				Sleep 1000
				
				; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지
				N41_ProcessingForPT.FromClickingPreAuthorizedButton_To_PrintOutPickTicket(CustomerPO, IsItFromNewOrder, IsItFromAllocation)
			}
		
			
		
		
		return
		}
	}







	; 오른쪽 위의 Open SO 화면의 픽티켓 버튼 클릭해서 인쇄하기
	; N41 의 pre authorized 버튼 클릭하지 않는다
	
	; 뉴오더에서 인쇄할 때 Order Status를 New Order 로 바꾸기 위해 표시하기 위한 IsItFromNewOrder
	OpenSO_Print_WITHOUT_PreAuthorized(CustomerPO, IsItFromNewOrder){
		
		; 1사분면의 Pick Ticket 버튼 누르기
		CommN41.ClickCreatePickTicketButton()

		
		; 눈에 거슬리지 않게 Gui 최소화 하기
		WinMinimize, Pick Ticket Processing
		WinMinimize, N41 Processing
		

		
		; Merge 확인 창이 나올지  Allocation 경고창이 나올지 모르기 때문에 일단 기다렸다 진행해야 됨	
		Sleep 4000
		
		; Merge 확인 창
		IfWinActive, SO Manager
		{
			Sleep 300
			Send, {Left}
			Sleep 500
			Send, {Left}
			Sleep 500
			Send, {Enter}
			Sleep 700
			
			WinWaitActive, SO Manager
			IfWinActive, SO Manager
			{
				Sleep 1000
				Send, {Enter}
				Sleep 1000			
				
				WinWaitActive, Pick Ticket ; Allocation 경고창
				IfWinActive, Pick Ticket
				{
					Sleep 1000
					Send, {Left}
					Sleep 500
					Send, {Enter}
					Sleep 700
					
					
					WinWaitActive, Pick Ticket ; Pick Ticket # 확인 창
					IfWinActive, Pick Ticket
					{
						Sleep 1000
						Send, {Enter}
						Sleep 800

						; pre authorized 버튼 클릭 않고 인쇄하기
						N41_ProcessingForPT.PrintWITHOUTClickPreAuthorizedButton(CustomerPO, IsItFromExcelFile, IsItFromNewOrder, IsItFromAllocation)
						return
					}				
				}			
			}
		}


		WinWaitActive, Pick Ticket ; Allocation 경고창
		;~ IfWinActive, Pick Ticket
		WinActivate, Pick Ticket
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
				
				Sleep 800
				; pre authorized 버튼 클릭 않고 인쇄하기
				N41_ProcessingForPT.PrintWITHOUTClickPreAuthorizedButton(CustomerPO, IsItFromExcelFile, IsItFromNewOrder, IsItFromAllocation)

			}
		}

		return
	}




















	; PreAuthorizedButton 누르는 것부터 프린트 하는 것까지
	FromClickingPreAuthorizedButton_To_PrintOutPickTicket(CustomerPO, IsItFromNewOrder, IsItFromAllocation){
		
			CN41_driver := New CommN41
		
			; pre authorized 버튼 클릭
			Text:="|<pre-authorize Button>*205$16.001zzbzyTztzzY0SE1tzzbzyTztzzc01zzy"
			if ok:=FindText(718,129,150000,150000,0,0,Text)
			{
				CoordMode, Mouse
				X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
				MouseMove, X+W//2, Y+H//2
				Click
				
				Sleep 2000
				
				; Pre-Authorized 통과 됐거나 Declined 됐을 때
				WinWaitActive Credit Card Processing, , 4
;				WinWaitActive Credit Card Processing
				IfWinActive, Credit Card Processing
				{
					Sleep 500
					WinActivate, Credit Card Processing					
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
								return
	;							Reload
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
							CommN41.ClickREfresh()
							
							Send, {Enter} ; 리프레쉬 버튼 누른 뒤 
							Sleep 2000
							
							result := CommN41.DoesThisPickTicketApproved() ; Approved 됐는지 화면에서 찾아본 뒤 찾았으면 1을 리턴하고 못 찾았으면 0을 리턴
							
							; 화면에서 Approved 못 찾았으면 주문 페이지 열기 원하는지 묻기
							if(result == 0){
								
								; Allocation 에서 호출된 것이 아닐때는 사용자에게 묻고 해당 주문 웹페이지 열기
								if(!IsItFromAllocation){
										
									MsgBox, 4100, NOT APPROVED, IT'S NOT CC APPROVED.`n`nWOULD YOU LIKE TO OPEN FG OR LAS PAGE OF THIS ORDER TO GET PRE-AUTHORIZATION ON THERE?
									IfMsgBox, Yes
									{
										
										WinClose, ahk_class Chrome_WidgetWin_1 ; 이상하게 열린 창 다시 사용할때 에러난다 차라리 다 닫고 새로 시작하자									
										
										FGFromAll_driver := New FG_ProcessingOfPT
										LASFromAll_driver := New LAS_ProcessingOfPT
										N_driver := new N41

										if(RegExMatch(CustomerPO, "imU)MTR")){
											IsItFromNewOrder = 0
											IsItFromExcelFile = 1
											FGFromAll_driver.ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) ; ##################### 주문 창 열기 #######################
										}			
										; CustomerPO 가 LAS 일때
										else if(RegExMatch(CustomerPO, "imU)OP")){
											LASFromAll_driver.ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) ; ##################### 주문 창 열기 #######################
										}
									}
								}
								; Allocation 에서 호출됐을때는 만약 카드가 Decline 나면 사람에게 묻지 말고 그냥 해당 주문 웹페이지 자동으로 열기
								else if(IsItFromAllocation){
										
									WinClose, ahk_class Chrome_WidgetWin_1 ; 이상하게 열린 창 다시 사용할때 에러난다 차라리 다 닫고 새로 시작하자									
										
									FGFromAll_driver := New FG_ProcessingOfPT
									LASFromAll_driver := New LAS_ProcessingOfPT
									N_driver := new N41

									if(RegExMatch(CustomerPO, "imU)MTR")){
										IsItFromNewOrder = 0
										IsItFromExcelFile = 1
										FGFromAll_driver.ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) ; ##################### 주문 창 열기 #######################
									}			
									; CustomerPO 가 LAS 일때
									else if(RegExMatch(CustomerPO, "imU)OP")){
										LASFromAll_driver.ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) ; ##################### 주문 창 열기 #######################
										
									}
								}
								
								
								; 쇼 오더가 아닐때 실행
								;~ if CustomerPO not contains 200
								;~ {

									SoundPlay, %A_WinDir%\Media\Ring06.wav
									
									
									;~ MsgBox, 266243, OPTIONS, THE PRE-AUTHORIZE BUTTON IS NOT CLICKED`nYOU CAN CHOOSE OPTIONS BELOW`n`n`nYes : CONTINUE THE PROCESSING`n`nNo : PUT "FG PA" IN THE HOUSE MEMO`n`nCancel : PUT "LAS PA" IN THE HOUSE MEMO
									MsgBox, 266243, OPTIONS, THE PRE-AUTHORIZE BUTTON IS NOT CLICKED`nYOU CAN CHOOSE OPTIONS BELOW`n`n`nYes : WRITE "FG PA" ON THE HOUSE MEMO`n`nNo : WRITE "LAS PA" ON THE HOUSE MEMO`n`nCancel : MOVE TO NEXT ORDER
									
									; 하우스 메모에 FG PA 입력하기
									IfMsgBox, Yes
									{
										CommN41_driver := New CommN41
										CommN41_driver.PutMemoIntoHouseMemoOnPickTicket()
										Send, FG PA
										Sleep 100
										Send, ^s
																				
										; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
										FG_PA = 1
										LAS_PA = 0
										N41_PA = 0
										CBS_or_ccDecline = 0								
										updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드										
										
;										MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, 메모 업데이트 됐는지 확인
										
									}
									; 하우스 메모에 LAS PA 입력하기
									IfMsgBox, No
									{
										CommN41_driver := New CommN41
										CommN41_driver.PutMemoIntoHouseMemoOnPickTicket()
										Send, LAS PA
										Sleep 100
										Send, ^s
										
										; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
										FG_PA = 0
										LAS_PA = 1
										N41_PA = 0
										CBS_or_ccDecline = 0								
										updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드										
										
;										MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, 메모 업데이트 됐는지 확인										
										
									}
									
									
									; CANCEL 눌렀을 때
									IfMsgBox, Cancel
									{
										; YES : CBS OR CC DECLINE
										; NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
										MsgBox, 4100, Wintitle, YES : CBS OR CC DECLINE`n`nNO : FIRST CC WAS DECLINE BUT OTHER CC WENT THROUGH
										{
											; YES : CBS OR CC DECLINE
											IfMsgBox, Yes
											{
												; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
												FG_PA = 0
												LAS_PA = 0
												N41_PA = 0
												CBS_or_ccDecline = 1
												updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드								
												
;												MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, CBS로 메모 업데이트 됐는지 확인
											} ; if ends - ; YES : CBS OR CC DECLINE
											
											; NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
											IfMsgBox, No
											{
												; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
												FG_PA = 0
												LAS_PA = 0
												N41_PA = 1
												CBS_or_ccDecline = 0
												updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드								
												
;												MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, N41 PA 로 메모 업데이트 됐는지 확인												
											
											} ; if ends - NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
										} ; if ends - 
									} ; if ends - CANCEL 눌렀을 때									
													
									


								;~ }

							}
							else if(result == 1){
									
								; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
								FG_PA = 0
								LAS_PA = 0
								N41_PA = 1
								CBS_or_ccDecline = 0								
								updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드
									
;								MsgBox, 262144, Approved Confirmed, 카드 통과됐음. 메모 저장됐는지 확인
							}

							
;							CommN41_driver := New CommN41
							
							; SO Manager 탭 클릭해서 pick ticket 탭에서 나오기
							;~ CommN41.OpenSOManager()
							CN41_driver.OpenSOManager()
							
							; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
							; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
							;~ CommN41.ClickREfreshButtonOnSOManager()
							CN41_driver.ClickREfreshButtonOnSOManager()
						

	;						Reload
							return
							
						}
	*/									
						
						
					}
								
					
				}
					
				; CC 가 없어서 업데이트 할거냐고 물을 때
				IfWinActive, Pick Ticket
				{
					WinActivate, Pick Ticket
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
	;							Reload
								return
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
							;~ MsgBox, 262144, Title, NO CC INFO ON THIS CUSTOMER`n`n`n`nCHECK THE SHIP VIA`n`nCLICK OK TO CONTINUE
							MsgBox, 4100, Wintitle, NO CC INFO ON THIS CUSTOMER`n`nWOULD YOU LIKE TO OPEN THE WEB PAGE OF THIS ORDER?
							IfMsgBox, Yes
							{			
								
								WinClose, ahk_class Chrome_WidgetWin_1 ; 이상하게 열린 창 다시 사용할때 에러난다 차라리 다 닫고 새로 시작하자									
								
								FGFromAll_driver := New FG_ProcessingOfPT
								LASFromAll_driver := New LAS_ProcessingOfPT
								N_driver := new N41

								if(RegExMatch(CustomerPO, "imU)MTR")){
									IsItFromNewOrder = 0
									IsItFromExcelFile = 1
									FGFromAll_driver.ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) ; ##################### 주문 창 열기 #######################
								}			
								; CustomerPO 가 LAS 일때
								else if(RegExMatch(CustomerPO, "imU)OP")){
									LASFromAll_driver.ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) ; ##################### 주문 창 열기 #######################
								}			

										
								;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
								;~ MsgBox, 266243, OPTIONS, THE PRE-AUTHORIZE BUTTON IS NOT CLICKED`nYOU CAN CHOOSE OPTIONS BELOW`n`n`nYes : CONTINUE THE PROCESSING`n`nNo : PUT "FG PA" IN THE HOUSE MEMO`n`nCancel : PUT "LAS PA" IN THE HOUSE MEMO
								MsgBox, 266243, OPTIONS, THE PRE-AUTHORIZE BUTTON IS NOT CLICKED`nYOU CAN CHOOSE OPTIONS BELOW`n`n`nYes : WRITE "FG PA" ON THE HOUSE MEMO`n`nNo : WRITE "LAS PA" ON THE HOUSE MEMO`n`nCancel : MOVE TO NEXT ORDER

								; 하우스 메모에 FG PA 입력하기
								IfMsgBox, Yes
								{
									CommN41_driver := New CommN41
									CommN41_driver.PutMemoIntoHouseMemoOnPickTicket()
									Send, FG PA
									Sleep 100
									Send, ^s
									
									; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
									FG_PA = 1
									LAS_PA = 0
									N41_PA = 0
									CBS_or_ccDecline = 0								
									updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드
																		
;									MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, 메모 업데이트 됐는지 확인									
									
								}
								; 하우스 메모에 LAS PA 입력하기
								IfMsgBox, No
								{
									CommN41_driver := New CommN41
									CommN41_driver.PutMemoIntoHouseMemoOnPickTicket()
									Send, LAS PA
									Sleep 100
									Send, ^s
									
									; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
									FG_PA = 0
									LAS_PA = 1
									N41_PA = 0
									CBS_or_ccDecline = 0								
									updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드
																		
;									MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, 메모 업데이트 됐는지 확인									
									
								}
								
								; CANCEL 눌렀을 때
								IfMsgBox, Cancel
								{
									; YES : CBS OR CC DECLINE
									; NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
									MsgBox, 4100, Wintitle, YES : CBS OR CC DECLINE`n`nNO : FIRST CC WAS DECLINE BUT OTHER CC WENT THROUGH
									{
										; YES : CBS OR CC DECLINE
										IfMsgBox, Yes
										{
											; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
											FG_PA = 0
											LAS_PA = 0
											N41_PA = 0
											CBS_or_ccDecline = 1
											updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드								
											
;											MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, CBS로 메모 업데이트 됐는지 확인
										} ; if ends - ; YES : CBS OR CC DECLINE
										
										; NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
										IfMsgBox, No
										{
											; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
											FG_PA = 0
											LAS_PA = 0
											N41_PA = 1
											CBS_or_ccDecline = 0
											updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드								
											
;											MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, N41 PA 로 메모 업데이트 됐는지 확인							
																				
										} ; if ends - NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
									} ; if ends - 
								} ; if ends - CANCEL 눌렀을 때		
								

							}

							; SO MANAGER 탭 누르고 끝내기
							CN41_driver := New CommN41
							CommN41.OpenSOManager()
							
							
							; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
							; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
							CommN41.ClickREfreshButtonOnSOManager()

							
	;						Reload
							return
							
						}
					}

					;~ Reload
					
				}

			}

		return
	} ; FromClickingPreAuthorizedButton_To_PrintOutPickTicket 메소드 끝
	
	
	



	
	
	
	; pre authorized 버튼 클릭 않고 인쇄하기
	PrintWITHOUTClickPreAuthorizedButton(CustomerPO, IsItFromExcelFile, IsItFromNewOrder, IsItFromAllocation){


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
	;				Reload
					return
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



				; Allocation 에서 온 주문이 아닐때는 주문 페이지를 열어볼 지 묻기				
				if(!IsItFromAllocation){
					
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 4100, Ask_To_Open_Web_Page, pre-authorized BUTTON IS NOT CLICKED`n`nWOULD YOU LIKE TO OPEN FG OR LAS PAGE OF THIS ORDER?
					
					IfMsgBox, Yes
					{
						WinClose, ahk_class Chrome_WidgetWin_1 ; 이상하게 열린 창 다시 사용할때 에러난다 차라리 다 닫고 새로 시작하자									
											
						FGFromAll_driver := New FG_ProcessingOfPT
						LASFromAll_driver := New LAS_ProcessingOfPT
						N_driver := new N41

						if(RegExMatch(CustomerPO, "imU)MTR")){
							IsItFromNewOrder = 0
							FGFromAll_driver.ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) ; ##################### 주문 창 열기 #######################
						}			
						; CustomerPO 가 LAS 일때
						else if(RegExMatch(CustomerPO, "imU)OP")){
							LASFromAll_driver.ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) ; ##################### 주문 창 열기 #######################
						}
					}
				}
				
				; Allocation 에서 온 주문은 그냥 주문 페이지 열기. (cbs보다 FG 혹은 LAS PA 얻어야 되는 일이 많으니까)
				else if(IsItFromAllocation){
					WinClose, ahk_class Chrome_WidgetWin_1 ; 이상하게 열린 창 다시 사용할때 에러난다 차라리 다 닫고 새로 시작하자									
											
					FGFromAll_driver := New FG_ProcessingOfPT
					LASFromAll_driver := New LAS_ProcessingOfPT
					N_driver := new N41

					if(RegExMatch(CustomerPO, "imU)MTR")){
						IsItFromNewOrder = 0
						FGFromAll_driver.ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) ; ##################### 주문 창 열기 #######################
					}			
					; CustomerPO 가 LAS 일때
					else if(RegExMatch(CustomerPO, "imU)OP")){
						LASFromAll_driver.ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) ; ##################### 주문 창 열기 #######################
					}				
				}

				
				;~ ; 쇼 오더가 아닐때 실행
				;~ if CustomerPO not contains 200
				;~ {

					;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
					;~ MsgBox, 266243, OPTIONS, THE PRE-AUTHORIZE BUTTON IS NOT CLICKED`nYOU CAN CHOOSE OPTIONS BELOW`n`n`nYes : CONTINUE THE PROCESSING`n`nNo : PUT "FG PA" IN THE HOUSE MEMO`n`nCancel : PUT "LAS PA" IN THE HOUSE MEMO
					MsgBox, 266243, OPTIONS, THE PRE-AUTHORIZE BUTTON IS NOT CLICKED`nYOU CAN CHOOSE OPTIONS BELOW`n`n`nYes : WRITE "FG PA" ON THE HOUSE MEMO`n`nNo : WRITE "LAS PA" ON THE HOUSE MEMO`n`nCancel : MOVE TO NEXT ORDER

					; 하우스 메모에 FG PA 입력하기
					IfMsgBox, Yes
					{
						CommN41_driver := New CommN41
						CommN41_driver.PutMemoIntoHouseMemoOnPickTicket()
						Send, FG PA
						Sleep 100
						Send, ^s
												
						; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
						FG_PA = 1
						LAS_PA = 0
						N41_PA = 0
						CBS_or_ccDecline = 0								
						updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드
															
;						MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, 메모 업데이트 됐는지 확인
						
					}
					; 하우스 메모에 LAS PA 입력하기
					IfMsgBox, No
					{
						CommN41_driver := New CommN41
						CommN41_driver.PutMemoIntoHouseMemoOnPickTicket()
						Send, LAS PA
						Sleep 100
						Send, ^s
						
						; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
						FG_PA = 0
						LAS_PA = 1
						N41_PA = 0
						CBS_or_ccDecline = 0								
						updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드
															
;						MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, 메모 업데이트 됐는지 확인						
						
					}
					
					; CANCEL 눌렀을 때
					IfMsgBox, Cancel
					{
						; YES : CBS OR CC DECLINE
						; NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
						MsgBox, 4100, Wintitle, YES : CBS OR CC DECLINE`n`nNO : FIRST CC WAS DECLINE BUT OTHER CC WENT THROUGH
						{
							; YES : CBS OR CC DECLINE
							IfMsgBox, Yes
							{
								; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
								FG_PA = 0
								LAS_PA = 0
								N41_PA = 0
								CBS_or_ccDecline = 1
								updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드								
								
;								MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, CBS로 메모 업데이트 됐는지 확인
							} ; if ends - ; YES : CBS OR CC DECLINE
							
							; NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
							IfMsgBox, No
							{
								; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
								FG_PA = 0
								LAS_PA = 0
								N41_PA = 1
								CBS_or_ccDecline = 0
								updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) ; 메소드								
								
;								MsgBox, 262144, PICK TICKET STATUS FILE UPDATED, N41 PA 로 메모 업데이트 됐는지 확인							}
							
							} ; if ends - NO  : 첫번째 카드 실패 후 다른 카드로 통과됐음
						} ; if ends - 
					} ; if ends - CANCEL 눌렀을 때

				;~ }


				CommN41_driver := New CommN41
				
				; SO MANAGER 탭 누르고 끝내기
				CN41_driver := New CommN41
				CN41_driver.OpenSOManager()

				; 아이템이 제대로 pick ticket에 들어갔는지 확인하기위해 SO Manager 에 있는 refresh 버튼 클릭해서
				; 가끔 store에 있는 정보가 다르면(예를 들어 52 street 과 52 st.) 아이템이 pick ticket에 안 들어가기도 한다
				CN41_driver.ClickREfreshButtonOnSOManager()
				
	;			Reload
				return
			}

		}	

		return
	} ; PrintWITHOUTClickPreAuthorizedButton 메소드 끝






























} ; N41_ProcessingForPT 클래스 끝
























































class LAS_ProcessingOfPT extends LA{


	; LASHOWROOM 오더 처리
	; 받은 PO 페이지 열어서 정보 읽고 UPDATE 버튼 누른 뒤 읽은 정보 리턴하기
	ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver){
		
		
				driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
				driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
				driver.AddArgument("--start-maximized") ; 창 최대화 하기			
		
				
				; LAS 페이지에서 정보 읽어서 저장하기
				; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서 
				; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
				Arr_LASInfo := processLAS_which_from_Not_New_Orders(driver, CustomerPO)


				Arr_BillingAdd := Arr_LASInfo[1].Clone()
				Arr_ShippingAdd := Arr_LASInfo[2].Clone()
				Arr_CC := Arr_LASInfo[3].Clone()
				Arr_Memo := Arr_LASInfo[4].Clone()
				ShippingStatus := Arr_LASInfo[5].Clone() ; shipping method 상태 저장 UPSG = 1	ETC = 2		LAS CONSOLIDATION = 3



				; UPSG 가 아니면 경고 메세지 띄우기
				if(ShippingStatus[1] == "1"){
	;				MsgBox, 262144, UPSG, It's UPSG
				}
				else if(ShippingStatus[1] == "2"){
					SoundPlay, %A_WinDir%\Media\Ring02.wav
					MsgBox, 262144, Title, It's neither UPSG nor LAS consolidation
				}
				else{
					SoundPlay, %A_WinDir%\Media\Ring02.wav
					MsgBox, 262144, LAS Consolidation, It's LAS Consolidation
				}

				
				BuyerNotes := Arr_Memo[1]
			;	AdditionalInfo := Arr_Memo[2] ; 이 정보는 없음
			;	StaffNotes := Arr_Memo[3] ; 이 정보는 없음
			;	CC# := Arr_CC[2] ; 이 정보는 없음


				; 필요 없는 문자가 들어있을 경우를 대비해 메모값 정리해주기
				BuyerNotes := Trim(BuyerNotes)
				BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")
				
				
				
				/* 배열로부터 읽기 첫 번째 방법
				Loop % Arr_ShippingAdd.Maxindex(){
					MsgBox % "Element number " . A_Index . " is " . Arr_ShippingAdd[A_Index]
				}
				*/			
							


				; 메모가 있을때만 창 키워서 표시하기
				;~ if(BuyerNotes || AdditionalInfo || StaffNotes){
				;~ if BuyerNotes not in None
				if(BuyerNotes)
				{				
					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
				}
				else{	

					SoundPlay, %A_WinDir%\Media\Ring06.wav
					MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
				}



				; No 눌렀으면 다시 시작
				IfMsgBox, No
				{
		;			Reload
					return
				}
				
				; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
				N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
				
				
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 262144, Title, Go to SO Manager Tab
				N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
				
				return
				

		
	} ; ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) 메소드 끝


























; ###############################################################################################################################################################################################################
; ###############################################################################################################################################################################################################
; 옛날 메소드 시작
; ###############################################################################################################################################################################################################
; ###############################################################################################################################################################################################################
	; LASHOWROOM 오더 처리
	00000000000000_OLD_000000000000000_ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver){

		; LASHOWROOM 오더 처리
		if(RegExMatch(CustomerPO, "imU)OP")){
			
			BuyerNotes := ""
			
			; LAS 페이지에서 정보 읽어서 저장하기
			; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서 
			; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
			Arr_FGInfo := LA.GetInfoFromLASPage(CustomerPO)


			Arr_BillingAdd := Arr_FGInfo[1].Clone()
			Arr_ShippingAdd := Arr_FGInfo[2].Clone()
			Arr_CC := Arr_FGInfo[3].Clone()
			Arr_Memo := Arr_FGInfo[4].Clone()

			
			BuyerNotes := Arr_Memo[1]
		;	AdditionalInfo := Arr_Memo[2] ; 이 정보는 없음
		;	StaffNotes := Arr_Memo[3] ; 이 정보는 없음
		;	CC# := Arr_CC[2] ; 이 정보는 없음


			; 필요 없는 문자가 들어있을 경우를 대비해 메모값 정리해주기
			BuyerNotes := Trim(BuyerNotes)
			BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")

			
			; 메모가 있을때만 창 키워서 표시하기
			;~ if(BuyerNotes || AdditionalInfo || StaffNotes){
			if BuyerNotes not in None
			{
				;~ SoundPlay, %A_WinDir%\Media\Ring02.wav ; Ring03 이 이상하면 Ring02 써보기
				SoundPlay, %A_WinDir%\Media\Ring03.wav
				MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}
			else{	

				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}



			; No 눌렀으면 다시 시작
			IfMsgBox, No
			{
	;			Reload
				return
			}
			
			; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
			N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
			
			; N41에 카드 정보가 있는지 확인하기 위해 카드 정보 입력창 열기
			N_driver.OpenRegisterCreditCard()

			Sleep 2000
			MsgBox, 262144, Memo, PLEASE CLICK Ok TO RESTART THE APPLICATION
			
			WinClose, Credit Card Management
			
			N_driver.ClickNewButtonOnCustomerMaster ; 끝내기 전에 뉴버튼 클릭하기
			N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기

	;		Reload
			return
		}

	} ; 	00000000000000_OLD_000000000000000_ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) 메소드 끝
; ###############################################################################################################################################################################################################
; ###############################################################################################################################################################################################################
; 옛날 메소드 끝
; ###############################################################################################################################################################################################################
; ###############################################################################################################################################################################################################









} ; LAS_ProcessingOfPT 클래스 끝























































class FG_ProcessingOfPT extends FG{


	; FG 오더 처리
	ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile){	
		
		
		N_driver := new N41	
	
		
		BuyerNotes := ""
		AdditionalInfo := ""
		StaffNotes := ""
		
		
		driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
		

		
		if(RegExMatch(CustomerPO, "imU)MTR")){
			
			
			; 전체 오더 검색창 주소로 이동하기
			URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
			driver := goToURl_AfterLogIn_IfNeeded(driver, URL) ; 원하는 url로 이동

;MsgBox, 전체 오더 검색창 화면으로 이동했음

			; 전체 오더 검색창 주소로 이동한 뒤
			; 검색조건을 PO 번호로 바꾼 뒤 PO 번호로 찾기
			driver := findOrdersByPO#(driver, CustomerPO)
			
;MsgBox, 검색 조건을 바꿨음

			; 가장 위에 있는 PO 번호를 새탭으로 열기
			driver := openNewTab_clickMostTopPO#(driver, CustomerPO)
			
;MsgBox, 새탭에서 열렸음

			; 현재 페이지의 Order Status 가 New Orders 이거나 Back Ordered 일때 Confirmed Orders 로 바꾸기
			driver := changeNewOrders_To_ConfirmedOrders(driver)
			
;MsgBox, Order Status 가 바뀌었음

			
			Arr_FGInfo := getInfoOnFG_And_Return_That(driver, CustomerPO, IsItFromNewOrder, IsItFromExcelFile)

;MsgBox, pause 1

			Arr_BillingAdd := Arr_FGInfo[1].Clone()
			Arr_ShippingAdd := Arr_FGInfo[2].Clone()
			Arr_CC := Arr_FGInfo[3].Clone()
			Arr_Memo := Arr_FGInfo[4].Clone()
			ShippingMethodStatus := Arr_FGInfo[5]

			
			BuyerNotes := Arr_Memo[1]
			AdditionalInfo := Arr_Memo[2]
			StaffNotes := Arr_Memo[3]
			CC# := Arr_CC[2]
			
			

			; 필요 없는 문자가 들어있을 경우를 대비해 메모들 값 정리해주기
			BuyerNotes := Trim(BuyerNotes)
			AdditionalInfo := Trim(AdditionalInfo)
			StaffNotes := Trim(StaffNotes)
			BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")
			AdditionalInfo := RegExReplace(AdditionalInfo, "[^a-zA-Z0-9 ]", "")
			StaffNotes := RegExReplace(StaffNotes, "[^a-zA-Z0-9 ]", "")
	
	
	

/*
for index, element in Arr_BillingAdd
{
	MsgBox % "Arr_BillingAdd " . index . " is " . element
}

for index, element in Arr_ShippingAdd
{
	MsgBox % "Arr_ShippingAdd " . index . " is " . element
}

for index, element in Arr_CC
{
	MsgBox % "Arr_CC 신용카드정보 " . index . " is " . element
}

for index, element in Arr_Memo
{
	MsgBox % "Arr_Memo " . index . " is " . element
}

MsgBox, % "ShippingMethodStatus : " . ShippingMethodStatus

MsgBox, 배열에 들어있는 값들 확인 완료
*/





			
			; UPS Ground 값은 3이다. 3이 아니면
			if(ShippingMethodStatus != 3)
			{
				SoundPlay, %A_WinDir%\Media\Ring02.wav
				MsgBox, 262144, UPS STATUS, IT IS NOT UPS GROUND SHIPMENT`n`nOK TO CONTINUE
			}


			; 고객정보 업데이트할지 묻지
			; 메모가 있을때만 창 키워서 표시하기
			if(BuyerNotes || AdditionalInfo || StaffNotes){
					
				SoundPlay, %A_WinDir%\Media\Ring02.wav ; Ring03 이 이상하면 Ring02 써보기
				;~ SoundPlay, %A_WinDir%\Media\Ring03.wav
				MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}
			; 메모 내용이 없으면 간단하게 업데이트 할지만 묻기
			else{

				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}

			; No 눌렀으면 고객정보 업데이트 하지 않기
			IfMsgBox, No
			{					
				; 뉴오더일때만 SO Manager 탭 열기
				; 뉴오더가 아니면 Pick Ticket 뽑다가 디클라인 난 뒤 웹페이지 호출했을 수 있으니 So Mangager 탭 여는게 더 귀찮기 때문					
				if(IsItFromNewOrder){
					N_driver.OpenSOManager() ; SO Manager 탭 열기
				}	
					
				return driver
	;			Reload
			}


			
			; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
			N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
			
			


			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, Title, Go to SO Manager Tab
			N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
			return driver
	;		Reload


		}
		
		return driver
	
	
	} ; ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) 메소드 끝


















	; FG 오더 처리
	0000000000_OLD_000000000000_ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile){
		
		BuyerNotes := ""
		AdditionalInfo := ""
		StaffNotes := ""
		

	/*
		; 크롬창이 열려있으면 닫을지 묻고 대답 없으면 5초 이내 닫기
		IfWinExist, ahk_class Chrome_WidgetWin_1
		{
			MsgBox, 4, , ALL CHROME BROWSERS WILL BE CLOSED.`nWOULD YOU LIKE TO CONTINUE?`n`nTHIS WINDOW IS CLOSED IN 5 SECONDS THEN CLOSE ALL CHROME WINDOWS, 3
			IfMsgBox No
			{
				MsgBox, 262144, WAIT TO CONTINUE, CLICK OK TO CONTINUE
			}
			
			WinClose, ahk_class Chrome_WidgetWin_1
			
		}
	*/	
		
		if(RegExMatch(CustomerPO, "imU)MTR")){
			
	;		MsgBox, % "CustomerPO : " . CustomerPO
			
			; FG 페이지에서 정보 읽어서 저장하기
			; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서
			; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
			Arr_FGInfo := FG.GettingInfoFromCurrentPage(CustomerPO, IsItFromNewOrder, IsItFromExcelFile)


			Arr_BillingAdd := Arr_FGInfo[1].Clone()
			Arr_ShippingAdd := Arr_FGInfo[2].Clone()
			Arr_CC := Arr_FGInfo[3].Clone()
			Arr_Memo := Arr_FGInfo[4].Clone()
			ShippingMethodStatus := Arr_FGInfo[5]

			
			BuyerNotes := Arr_Memo[1]
			AdditionalInfo := Arr_Memo[2]
			StaffNotes := Arr_Memo[3]
			CC# := Arr_CC[2]
			
			

			; 필요 없는 문자가 들어있을 경우를 대비해 메모들 값 정리해주기
			BuyerNotes := Trim(BuyerNotes)
			AdditionalInfo := Trim(AdditionalInfo)
			StaffNotes := Trim(StaffNotes)
			BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")
			AdditionalInfo := RegExReplace(AdditionalInfo, "[^a-zA-Z0-9 ]", "")
			StaffNotes := RegExReplace(StaffNotes, "[^a-zA-Z0-9 ]", "")
						
			
			; UPS Ground 값은 3이다. 3이 아니면 
	;		MsgBox, % ShippingMethodStatus
			if(ShippingMethodStatus != 3)
			{
				SoundPlay, %A_WinDir%\Media\Ring02.wav
				MsgBox, 262144, UPS STATUS, IT IS NOT UPS GROUND SHIPMENT`n`nOK TO CONTINUE
			}


			; 고객정보 업데이트할지 묻지
			; 메모가 있을때만 창 키워서 표시하기
			if(BuyerNotes || AdditionalInfo || StaffNotes){
					
				SoundPlay, %A_WinDir%\Media\Ring02.wav ; Ring03 이 이상하면 Ring02 써보기
				;~ SoundPlay, %A_WinDir%\Media\Ring03.wav
				MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}
			; 메모 내용이 없으면 간단하게 업데이트 할지만 묻기
			else{

				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}

			; No 눌렀으면 고객정보 업데이트 하지 않기
			IfMsgBox, No
			{					
				; 뉴오더일때만 SO Manager 탭 열기
				; 뉴오더가 아니면 Pick Ticket 뽑다가 디클라인 난 뒤 웹페이지 호출했을 수 있으니 So Mangager 탭 여는게 더 귀찮기 때문					
				if(IsItFromNewOrder){
					N_driver.OpenSOManager() ; SO Manager 탭 열기
				}	
					
				return
	;			Reload
			}


			
			; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
			N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
			
			


			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, Title, Go to SO Manager Tab
			N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
			return
	;		Reload


		}
		
		return
		
	} ; ProcessingFGOrder 메소드 끝











} ; FG_Processing 클래스 끝




















































class Allocation extends CommEXCEL{

	ProcessingAllocationList(){

		; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		IfWinNotExist, ahk_class XLMAIN
		{
			Loop{
				MsgBox, 262144, No Excel file Warning, PLEASE OPEN THE ALLOCATION EXCEL FILE.
					
				IfWinExist, ahk_class XLMAIN
					break
			}
		}
		
			
		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible
		
		return
		
	} ; ProcessingAllocationList; 메소드 끝
	
	
	
	
	
	
	
	
} ; Allocation 클래스 끝




























































class BO extends CommEXCEL{


	; Style -> Style SO Detail 에 들어가서 엑셀 파일 만든다
	; 창이 열려있는 상태에서 SO # 읽는다
	; 메소드 호출한다
	; 끝까지 진행해서 번호가 없으면 프로그램 리로드 한다
	BOProcess_by_Style(IsItFromAllocation){
		
;		MsgBox, method in
			
		; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
		IfWinNotExist, ahk_class XLMAIN
		{
			Loop{
				MsgBox, 262144, No Excel file Warning, PLEASE OPEN BO LIST EXCEL FILE
					
				IfWinExist, ahk_class XLMAIN
					break
			}
		}

			
			
		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible


		; 엑셀에 값이 들어간 만큼(i 값 만큼) 루프 돌면서 엑셀에서 값 읽기
		Loop{

			; SO# 값은 B Column 두 번째에 에 있음
			SO# := Xl.Range("B2").Value
			
			
			; Allocation 파일을 읽는 호출일 때(IsItFromAllocation 변수에 값이 있을 때) 실행
			; 엑셀에서 SO# 읽어서 변수에 저장
			if(IsItFromAllocation){
				; SO# 값은 A Column 첫 번째에 에 있음
				SO# := Xl.Range("A1").Value				
			}
			
			;소수점 뒷자리 정리
			RegExMatch(SO#, "imU)(\d*)\.", SubPat)
				
			; 정리된 값 SO# 에 넣기
			SO# := SubPat1			
			
	

			; 만약 지금 얻은 SO# 값이 이전 SO# 값을 저장하고 있는 previousNumber 값과 같다면 
			; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
			; 그 전에 SO# 변수와 previousNumber 변수에 값이 있을때만 비교하기
			if(SO# || previousNumber){
				
				IfEqual, SO#, %previousNumber%
				{
					Xl.Sheets(1).Range("A2").EntireRow.Delete
					continue
				}				
			}

			

;			MsgBox, % SO#
			
			; 해당 SO# 검색하고 처리하는 메소드 실행
			; StartSO# 값의 자리에 SO#값을 넣어서 호출하기
			; 여기서 호출했다는 것을 표시하기 위해 IsItFromExcelFile 변수 1로 만들어서 호출
			; 뉴오더에서 호출된 것이 아니니까 IsItFromNewOrder 값은 0으로 만들기
			IsItFromExcelFile = 1
			IsItFromNewOrder = 0
			SO#FromJODIFLcom = 0
			BasicProcessing(SO#, EndSO#, CustomerPO, IsItFromNewOrder, IsItFromExcelFile, SO#FromJODIFLcom, CustMemoFromJODIFLcom, IsItFromAllocation)
			
;MsgBox, 262144, Title, BasicProcessing 함수에서 돌아왔음 1486

			
			; 값이 없으면 루프 나오면서 끝내기
			if(!SO#){
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 262144, END OF THE LIST, END OF THE LIST		
					
				; 엑셀 종료하기
;				xl.WorkBooks.Close()                                    ;close file
;				xl.quit
				
				Reload
	;			break
			}


			; 백오더 처리일때 실행
			; Allocation 파일을 읽는 호출이 아닐때 때(IsItFromAllocation 변수에 값이 없을 때) 실행
			if(!IsItFromAllocation){

				; 두 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
				Xl.Sheets(1).Range("A2").EntireRow.Delete


				; 중복되는 값의 비교를 위해 previousNumber 변수에 SO# 값 넣기
				previousNumber := SO#
			}
			

			
			; Allocation 파일을 읽는 호출일 때(IsItFromAllocation 변수에 값이 있을 때) 실행			
			if(IsItFromAllocation){

				; 첫 번째 Row 값은 읽어서 오더를 처리했으니 엑셀에서 지워주기
				Xl.Sheets(1).Range("A1").EntireRow.Delete
				
				; previousNumber 변수에 항상 같은 값을 집어넣어서 위의 IfEqual, SO#, %previousNumber% 코드가 실행되지 않게 하기
				previousNumber = 1
			}
			
			; 크롬 창 닫기
			WinClose, ahk_class Chrome_WidgetWin_1

		}
		
		return
		
	} ; BOProcess_by_Style 메소드 끝
	
	
	
	
	
	
	
	
	
	
	
	
} ; BO 클래스 끝

























































class JODIFL extends CommWeb{

	ProcessingJodiflcom(){
		
;		MsgBox, in the ProcessingJodiflcom Method
		
		; 변수들 값 초기화
		SO#ofThisWebOrder := ""
		ShipmentMethod := ""
		CustMemo := ""
		CustPo#OfJodiflWeb := ""
		

		
		; 파일에서 읽기. LastCustPo#OfJodiflWeb.txt 파일의 내용을 CustPo#OfJodiflWeb 변수에 저장하기
		; CustPo#OfJodiflWeb 변수에는 이번에 열어볼 Web: 으로 시작되는 jodifl.com 의 Customer PO 번호가 들어가 있음
		; 이 값 사용해서 해당 Customer PO 의 Url 을 열 것임
		FileRead, CustPo#OfJodiflWeb, %A_ScriptDir%\CreatedFiles\jodifl.com\LastCustPo#OfJodiflWeb.txt
		
		
		; 이전 오더 페이지 번호 저장된 파일에서 값 읽어서 PreviouseUrl 변수에 넣기
		FileRead, PreviouseUrl, %A_ScriptDir%\CreatedFiles\jodifl.com\PreviouseUrl.txt

;		MsgBox, % CustPo#OfJodiflWeb
		
		
		; 오더 페이지 열기 위한 기본 url
		; 여기에 개별 Customer PO 번호를 더해서 해당 오더 페이지를 열게된다
		BasicUrlOfEachOrderPages = https://www.jodifl.com/index.php/admin/sales_order/view/order_id/
		
		
		
		; 이전에 처리된 JODIFL 창이 없으면 새로 창 열기
		; 일단 현재 크롬창이 열려 있는지 확인 하기
		;~ IfWinExist, ahk_class Chrome_WidgetWin_1
		IfWinExist, ahk_exe chrome.exe
		{
;			MsgBox, 크롬창 열려있음
			driver := ChromeGet()
;			MsgBox, % driver.Window.Title

		
			; 지금 열려있는 크롬창의 WINDOW TITLE 에 Magento Admin 가 포함되어 있지 않으면 JODIFL 창 새로 열기
			if driver.Window.Title not contains Magento Admin
			{
				; JODIFL 창 열기	
				CommWeb.OpenNewBrowser(BasicUrlOfEachOrderPages)
				Sleep 500
				driver := ChromeGet()
			}

		}
		else ; 현재 크롬창이 열려있지 않으면 JODIFL 창 새로 열기
		{
			CommWeb.OpenNewBrowser(BasicUrlOfEachOrderPages)
			driver := ChromeGet()
		}

/*		
		; 이전에 처리된 창이 열려있지 않으면 새로 창 열기
		; 이전에 처리된 창이 열려있으면 그냥 그 창 사용하기
		NamePreviousGoogleBrowser := "#100000" . CustPo#OfJodiflWeb . " / Orders / Sales / Magento Admin - Google Chrome"
MsgBox, % "WINDOW TITLE OF THIS PAGE IS : " . NamePreviousGoogleBrowser
		IfWinNotExist, %NamePreviousGoogleBrowser%
		{			
			; JODIFL 창 열기			
			CommWeb.OpenNewBrowser(BasicUrlOfEachOrderPages)
			driver := ChromeGet()
		}		
*/
		
		
		



		; 현재 페이지가 로그인 페이지면 자동 로그인 후 재귀호출 하기	

		; 로그인 페이지의 ID, 패스워드 입력칸 Xpath
		ID_Xpath = //*[@id='username']
		Password_Xpath = //*[@id='login']
		
		ID = admin
		PW = Jodifl1234
		
		
		
		; 만약 현재 페이지에 패스워드 입력칸 Xpath 가 나타나 있다면 이건 로그인 페이지라는 뜻
		if(driver.FindElementByXPath(Password_Xpath))
		{			
			
			; ID 입력 후 엔터, 비밀번호 입력 후 엔터치기
			driver.FindElementByXPath(ID_Xpath).SendKeys(ID).sendKeys(driver.Keys.ENTER)
			driver.FindElementByXPath(Password_Xpath).SendKeys(PW).sendKeys(driver.Keys.ENTER)
			
			driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다

			
/*			
			; 이미 크롬에 저장된 비밀번호 사용해서 편법으로 로그인하기			
			; 패스워드 입력칸으로 이동한 뒤
			driver.FindElementByXPath(Password_Xpath).click()			

			; 그곳에서 화살표 내려서 이미 저장되어 있는 비밀번호 선택 하기
			Send, {Down}
			Sleep 100
			send, {Enter}
			Sleep 100
			send, {Enter}
*/

			; 재귀호출로 메소드 다시 시작하기			
			JODIFL.ProcessingJodiflcom()
		}		
		
		
		; LastCustPo#OfJodiflWeb.txt 파일에는 이전에 처리된 Customer PO 번호가 있기 때문에 1을 증가해주기
		CustPo#OfJodiflWeb++
				
		
		
		; URL을 Customer PO 에 맞게 고치기
		; 맨 끝에 Customer PO 번호 붙이기
		URLAddedCustPO := BasicUrlOfEachOrderPages . CustPo#OfJodiflWeb
		
		
;		MsgBox, % "OPEN THE URL`n" . URLAddedCustPO
		
		; 해당하는 오더 페이지 열기
		driver.Get(URLAddedCustPO)
		
		



		; 현재 url 얻기
		driver.executeScript("return document.readyState").toString().equals("complete") ; 페이지가 로딩이 끝날때까지 기다립니다
		sleep 700
		CurrentURL := driver.Url
		
		
		
		; 현재 페이지가 이전 페이지와 같으면 재귀호출로 다시 시작하기
		if(CurrentURL == PreviouseUrl){
;			SoundPlay, %A_WinDir%\Media\Ring02.wav
			;~ MsgBox, 262144, warning, IT DOESNT OPEN NEW ORDER PAGE EITHER NO NEW ORDER RECEIVED SINCE THE LAST PROCESSED ORDER
			MsgBox, 262144, warning, IT DOESNT OPEN NEW ORDER PAGE
;			driver.refresh()
			JODIFL.ProcessingJodiflcom()
		}
		

		
		; 현재 페이지가 전체화면 페이지면 Customer PO 를 끝까지 처리해서 처리할 것이 없다는 뜻이므로		
		; 마무리하고 프로그램 끝내기 (다시 시작하기)
		
		; 전체화면 url		
		URLofDashBoardOfJodiflcom = https://www.jodifl.com/index.php/admin/sales_order/index/key/		
		
		if CurrentURL contains %URLofDashBoardOfJodiflcom%
		{
			
			MsgBox, 262144, End of Order List, NO MORE ORDERS`n`nOK TO RESTART THE APPLICATION
			
			; 위에서 마지막 처리된 jodifl.com 의 Customer PO 를 1 증가시켰을 때 그런 오더가 없다고 전체 화면이 되었으므로 변수에서 1을 빼준뒤 파일에 저장. 이렇게 1을 빼줌으로서 다음에 이 메소드 호출될때 마지막으로 처리된 번호부터 시작될 것임
			CustPo#OfJodiflWeb--

			; LastCustPo#OfJodiflWeb.txt 파일 내용 덮어쓰기(초기화) 위해 파일 지우기
			FileDelete, %A_ScriptDir%\CreatedFiles\jodifl.com\LastCustPo#OfJodiflWeb.txt		
						
			; 파일에 쓰기. CustPo#OfJodiflWeb 변수 안에 있는 내용 LastCustPo#OfJodiflWeb.txt 파일에 쓰기
			FileAppend, %CustPo#OfJodiflWeb%, %A_ScriptDir%\CreatedFiles\jodifl.com\LastCustPo#OfJodiflWeb.txt
			
			Reload
		}



		; 화면 밑의 Sales Order 번호, 고객 메모 등 있는 곳의 Xpath
		BasicInfo_Xpath = //*[@id='order_history_block']/ul/li
		
		Loop{
			
			; 어떤건 Xpath 하나로 찾을 수 있는 것도 있고 어떤건 그 다음도 찾아야 되는 것이 있어서 기본 BasicInfo_Xpath 에 [1], [2] ... 이런식으로 더해가면서 값을 찾기 위해			
			Info_Xpath := BasicInfo_Xpath . "[" . A_Index . "]"
			
;			MsgBox, % Info_Xpath
			
			; 화면 밑의 SO# 찾기위해 해당 Xpath 에 들어있는 값 InfoInSO#ofThisWebOrder 변수에 넣기
			InfoInSO#ofThisWebOrder := driver.FindElementByXPath(Info_Xpath).Attribute("textContent")


			; 읽은 값중에 Processing 이 들어있으면 값 없애기
;			if(RegExMatch(InfoInSO#ofThisWebOrder, "Processing"))
;				InfoInSO#ofThisWebOrder := ""


			; 읽은 값이 Pending 이 아니면 값 없애기
			if(!RegExMatch(InfoInSO#ofThisWebOrder, "Pending"))
				InfoInSO#ofThisWebOrder := ""


			; Sales Order 번호가 없으면 번호 찾아서 변수에 넣기. 있으면 하지 않기
			; 한개의 Xpath에 모든 값이 들어있는 것이 아니다. 때문에 Info_Xpath 에 있는 Xpath 값을 하나씩 올려가면서 찾게 된다
			; 먼저 Sales Order 번호를 찾고 ShipmentMethod 와 메모는 그 다음 위치에 있다
			; 만약 아래 if 문을 쓰지 않는다면 이미 앞에서 찾은 Sales Order 번호가 저장된 SO#ofThisWebOrder 변수에 이상한 값을 덮어씌우게 된다.
			; 그렇기 때문에 SO#ofThisWebOrder 변수에 값이 있다면 지금 보고있는 Xpath 값 이전에 이미 Sales Order 번호를 찾았다는 뜻이므로 그냥 넘어가게끔 if를 사용해서 SO#ofThisWebOrder 변수에 값을 넣어준다
			if(!SO#ofThisWebOrder){
				SO#ofThisWebOrder := RegExReplace(InfoInSO#ofThisWebOrder, ".*N41\sSO#\s(\d)", "$1")  ; $1 역참조를 사용하여 N41 SO# 다음의 숫자만 돌려줍니다				
			}
			
			
			; Shipment Method 어떤지 찾기
			; if 문을 사용한 이유는 위의 SO#ofThisWebOrder 에 if 문을 사용한 이유와 같다. 기껏 찾은 값에 다른 값을 덧씌우지 않기 위해
			if(!ShipmentMethod){					
				RegExMatch(InfoInSO#ofThisWebOrder, "imU)(Via\](.*)\[Ship)", SubPat)
				ShipmentMethod := SubPat2
				driver.FindElementByXPath(Info_Xpath).click()
				
;				MsgBox, % "ShipmentMethod`n" . ShipmentMethod
			}


			
			; 고객 메모 있는지 찾기
			; if문을 사용한 이유는 위와 같음
			if(!CustMemo){
				RegExMatch(InfoInSO#ofThisWebOrder, "(Via\].*Memo](.*))", SubPat)
				CustMemo := SubPat2
				
;				MsgBox, % "CustMemo`n" . CustMemo
			}
			
			
			; Sales Order 번호와 Shipment Method 를 찾았으면 루프 끝내고 값을 찾은 변수들 리턴하기
			; CustMemo 까지 확인하지 않는 건 메모를 남기지 않은 고객도 있을 수 있기 때문
			if(SO#ofThisWebOrder && ShipmentMethod){

				; PreviouseUrl.txt 파일 내용 덮어쓰기(초기화) 위해 파일 지우기		
				FileDelete, %A_ScriptDir%\CreatedFiles\jodifl.com\PreviouseUrl.txt		
				
				; 파일에 쓰기. CustPo#OfJodiflWeb 변수 안에 있는 내용 LastCustPo#OfJodiflWeb.txt 파일에 쓰기
				FileAppend, %URLAddedCustPO%, %A_ScriptDir%\CreatedFiles\jodifl.com\PreviouseUrl.txt
				
				
				
				; LastCustPo#OfJodiflWeb.txt 파일 내용 덮어쓰기(초기화) 위해 파일 지우기
				FileDelete, %A_ScriptDir%\CreatedFiles\jodifl.com\LastCustPo#OfJodiflWeb.txt		
				
				; 파일에 쓰기. CustPo#OfJodiflWeb 변수 안에 있는 내용 LastCustPo#OfJodiflWeb.txt 파일에 쓰기
				FileAppend, %CustPo#OfJodiflWeb%, %A_ScriptDir%\CreatedFiles\jodifl.com\LastCustPo#OfJodiflWeb.txt
				
;				MsgBox, % "SO#`n" . SO#ofThisWebOrder . "`n`nShipmentMethod`n" . ShipmentMethod . "`n`nCustMemo`n" . CustMemo
				return [SO#ofThisWebOrder, ShipmentMethod, CustMemo]
			}

		} ; 루프 끝


		;~ Info_Xpath = //*[@id='order_history_block']/ul/li[1]
		;~ Info_Xpath = //*[@id='order_history_block']/ul/li[2]
		
		
		
		
		MsgBox, % InfoInSO#ofThisWebOrder
		
		SO#ofThisWebOrder := RegExReplace("abcXYZ123", "abc(.*)123", "aaa$1zzz")  ; $1 역참조를 사용하여 "aaaXYZzzz"을 돌려 줍니다.
		SO#ofThisWebOrder := RegExReplace(InfoInSO#ofThisWebOrder, "abc(.*)123", "$1")  ; $1 역참조를 사용하여 "XYZ"을 돌려 줍니다.
		SO#ofThisWebOrder := RegExReplace(InfoInSO#ofThisWebOrder, ".*N41\sSO#\s(\d)", "$1")  ; $1 역참조를 사용하여 "XYZ"을 돌려 줍니다.
		
		MsgBox, % SO#ofThisWebOrder
		
		
		
		
		
		
		RegExMatch(SO#ofThisWebOrder, "imU)MTR")
		
		
		
		
		
		
		
		MsgBox, OPNED?






		
		; 파일 내용 덮어쓰기 위해 파일 지우기		
		FileDelete, %A_ScriptDir%\CreatedFiles\jodifl.com\LastCustPo#OfJodiflWeb.txt		
		
		; 파일에 쓰기. CustPo#OfJodiflWeb 변수 안에 있는 내용 LastCustPo#OfJodiflWeb.txt 파일에 쓰기
		FileAppend, %CustPo#OfJodiflWeb%, %A_ScriptDir%\CreatedFiles\jodifl.com\LastCustPo#OfJodiflWeb.txt
		
		
		MsgBox
		
		
		




			
		return
		
	} ; ProcessingJodiflcom 메소드 끝



} ; JODIFL 클래스 끝














; 현재 화면 pick ticket 의 상태(FG PA, LAS PA, N41 PA, CBS)를 파일에 저장하는 메소드
updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline){
	
;	MsgBox, 262144, Title, updPTStatus(FG_PA, LAS_PA, N41_PA, CBS_or_ccDecline) 메소드 들어왔음
	

	; 오늘 날짜 변수에 저장
	TodaysDate := A_YYYY . A_MM . A_DD
;	MsgBox, % TodaysDate

	; 오늘 날짜 파일(pick # 등 저장할 파일)이 없으면 해당 폴더에 만들기
	IfNotExist, %A_ScriptDir%\CreatedFiles\P.T. Status\PT_CH_%TodaysDate%.txt
	{    
		FileAppend, , %A_ScriptDir%\CreatedFiles\P.T. Status\PT_CH_%TodaysDate%.txt
		MsgBox, 262144, CREATE A FILE, 오늘 날짜 파일이 없어서 만들었습니다. 계속 진행합니다
	}


	; pick ticket 화면에서 pick#, 고객코드, 만든날짜 읽은 후 리턴

	CommN41_driver := New CommN41
	infoOnPickTicket_Array := object()
	infoOnPickTicket_Array := CommN41_driver.getInfoOnPickTicket_Then_ReturnThem()
	
/*	
	infoOnPickTicket_Array := object()
	infoOnPickTicket_Array := getInfoOnPickTicket_Then_ReturnThem()	
*/	
	


	pick# := infoOnPickTicket_Array[1]
	custCode := infoOnPickTicket_Array[2]
	updDate := infoOnPickTicket_Array[3]
	
	
;	MsgBox, 262144, Title, 리턴받은 값들 확인%pick#%`n`n`n%custCode%`n`n`n%updDate%

	;~ MsgBox, % "infoOnPickTicket_Array[1] : " . infoOnPickTicket_Array[1] . "`n`ninfoOnPickTicket_Array[2] : " . infoOnPickTicket_Array[2] . "`n`ninfoOnPickTicket_Array[3] : " . infoOnPickTicket_Array[3]
	;`n`n`n`n`n`n`ninfoOnPickTicket_Array[1] : %infoOnPickTicket_Array[1]%`ninfoOnPickTicket_Array[2] : %infoOnPickTicket_Array[2]%`n`ninfoOnPickTicket_Array[3] : %infoOnPickTicket_Array[3]%


	; pick ticket 화면에 있는 pick ticket 번호, 고객코드, 만든날짜 를 파일명이 오늘 날짜인 파일에 저장하기 i.g.) 20180802.txt 파일에 저장
	; 세 변수에 모두 값이 있을때만 동작
	if(pick# && CustCode && updDate){		
		
		if(FG_PA){
			FileAppend, %pick#%`t%custCode%`t%updDate%`tFG PA`n, %A_ScriptDir%\CreatedFiles\P.T. Status\PT_CH_%TodaysDate%.txt
		}
		else if(LAS_PA){
			FileAppend, %pick#%`t%custCode%`t%updDate%`tLAS PA`n, %A_ScriptDir%\CreatedFiles\P.T. Status\PT_CH_%TodaysDate%.txt
		}
		else if(N41_PA){
			FileAppend, %pick#%`t%custCode%`t%updDate%`tN41 PA`n, %A_ScriptDir%\CreatedFiles\P.T. Status\PT_CH_%TodaysDate%.txt
		}
		else if(CBS_or_ccDecline){
			FileAppend, %pick#%`t%custCode%`t%updDate%`tCBS`n, %A_ScriptDir%\CreatedFiles\P.T. Status\PT_CH_%TodaysDate%.txt
		}
	}
	
	; 변수들 중 한 개에라도 값이 아무것도 없으면 다시 시작
	else if(!pick# || !CustCode || !updDate){
		MsgBox, 262144, Title, 변수들 중에 값이 없는 변수가 있습니다. 파일 저장 동작을 건너뜁니다. 이 메세지 창을 닫기 전에 파일을 수동으로 업데이트 하시기 바랍니다.`n`n`n리턴받은 값들 확인`n`n%pick#%`n`n`n%custCode%`n`n`n%updDate%
	}	

		
	return
		
	
} ; updPTStatus 메소드 끝









