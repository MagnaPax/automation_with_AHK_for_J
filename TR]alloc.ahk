; ######################################################################
; Allocation 파일 보기 편하게 쓸데 없는 값 지운 뒤 중복되는 SO 번호 지우기
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




	MsgBox, 262144, REFINE ALLOCATION FILE, CLICK OK TO CONTINUE


	; 화면 모드 Relative로 설정하기
	CoordMode, Mouse, Relative

	; 사용자의 키보드와 마우스 입력은 Click, MouseMove, MouseClick, 또는 MouseClickDrag이 진행 중일 때 무시됩니다 
	BlockInput, Mouse



	; 만약 엑셀 창이 열려있지 않으면 열릴때까지 무한 반복으로 경고창 표시하기
	IfWinNotExist, ahk_class XLMAIN
	{
		loop{
			MsgBox, 262144, No Excel file Warning, PLEASE OPEN AN ORDER EXCEL FILE
			IfWinExist, ahk_class XLMAIN
				break
		}
	}
	
	
	
	; 아이템 찾는 동안 보여줄 프로그래스 바 
	TotalLoops = 20000
	Gui, -Caption +AlwaysOnTop +LastFound
	Gui, Add, Text, x12 y9 w170 h20 , P  R  O  C  E  S  S  I  N  G  .  .  .
	Gui, Add, Progress, w410 Range0-%TotalLoops% vProgress
	Gui, Show, w437 h84, SEARCHING ITEMS




	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible		
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
	

	
	; 1번째 줄(Row) 지우기
	Xl.Sheets(1).Range("A1").EntireRow.Delete
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
	

	;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
	XL_Handle(XL,1) ;get handle to Excel Application
	;~ i := XL_Last_Row(XL)
	lastLine#OfRow := XL_Last_Row(XL)
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
	
;	MsgBox, % i
	
	
	; A~B, D, F, H~X 열(Columns) 지우기
	;~ XL_Col_Delete(XL,RG:="A:B|D|F|H:X") ;Delete columns	
	
		

	; 열(Columns) 지우기
	XL_Col_Delete(XL,RG:="A:B|D:E|G:AK") ;Delete columns
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	
	

	; 열의 넓이 설정하기
	XL_Col_Width_Set(XL,RG:="A=10|B=25")
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
	
	
	; 스타일 번호와 색깔이 들어있는 Row 번호
	#ofLineOfStyle#AndColor = 1
	Loop{
		
		; 지울 셀(스타일 번호와 색깔이 들어있는 Row의 A 셀)
		#ofDeleteRow := "A" . #ofLineOfStyle#AndColor
	
		; 스타일 번호 들어있는 셀 위치
		cell#OfStyle#ToBeCopied := "A" . #ofLineOfStyle#AndColor
		
		; 그 스타일 색깔이 들어있는 셀 위치
		cell#OfColorOfTheStyle#ToBeCopied := "B" . #ofLineOfStyle#AndColor
		
		
		; Style 번호와 색깔 읽어와서 각각 변수에 저장
		Style# := Xl.Range(cell#OfStyle#ToBeCopied).Value
		Color := Xl.Range(cell#OfColorOfTheStyle#ToBeCopied).Value
		
		
;		MsgBox, % Style# . "`n`n" . Color
		
		; 주문번호와 고객 코드가 있는 줄(스타일 번호와 색깔이 저장될 Row 번호)
		;~ #ofLineOfSO#AndCustomerCode := #ofLineOfStyle#AndColor + 1		

		; 스타일 번호가 저장될 셀 위치(주문 번호와 회사명이 있는 Row 옆에 위치시키기)
		cell#ToBeStoredTheStyle# := "C" . #ofLineOfStyle#AndColor + 1
		
		; 그 스타일 색깔이 저장될 셀 위치(주문 번호와 회사명이 있는 Row 옆에 위치시키기)
		cell#ToBeStoredTheColorOfTheStyle# := "D" . #ofLineOfStyle#AndColor + 1
		
		
		; 새로운 위치에 스타일 번호와 색깔 저장
		Xl.Range(cell#ToBeStoredTheStyle#).Value := Style#
		Xl.Range(cell#ToBeStoredTheColorOfTheStyle#).Value := Color
		
		
;		MsgBox, 262144, Title, 복사됐음
		
		
		; 스타일 번호와 색깔이 있는 홀수 줄 지우기
		Xl.Sheets(1).Range(#ofDeleteRow).EntireRow.Delete
		

		;
		#ofLineOfStyle#AndColor := #ofLineOfStyle#AndColor + 1
		
		
		; 마지막에 도착하면 루프 끝내고 나가기
		if(#ofLineOfStyle#AndColor > lastLine#OfRow)
			break		
		

		GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
		
	}





; #############################################################
;	MsgBox, 262144, Title, 고객명(2열)로 정렬 후 중복값 지우기
; #############################################################	
	
	; 중복값 지우기 위해 일단
	; 2 열(Columns)을 정렬하기
	xl.cells.sort(xl.columns(2), 1)
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	



	;엑셀 값의 끝 row 번호 알아낸 후 i 에 값 넣기
	XL_Handle(XL,1) ;get handle to Excel Application
	i := XL_Last_Row(XL)
	GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가	


;~ MsgBox
		
		j = 1
		Loop{

			k := j + 1

			#ofRowToRead := "B" . j
			#ofRowToReadAddedOne :=  "B" . k
			
			
			var1 := Xl.Range(#ofRowToRead).Value
			var2 := Xl.Range(#ofRowToReadAddedOne).Value
			
;			MsgBox, % var1 . "`n" . var2
			
			
			if(var1 == ""){
				break
			}
						

			; 만약 지금 얻은 SO# 값이 이전 SO# 값을 저장하고 있는 previousNumber 값과 같다면 
			; 중복된 값이니 현재 Row 삭제한 뒤 루프 처음으로 돌아가기
			IfEqual, var1, %var2%
			{
				
;				MsgBox, % var1 . "`n" . var2 . "`n`n" . "delete" . "`n`n" . "i : " . i . "`nj : " . j
				Xl.Sheets(1).Range(#ofRowToReadAddedOne).EntireRow.Delete
				
				GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가				
				continue
			}
			
			IfNotEqual, var1, %var2%
			{
				j++
				
				GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가
				continue
			}
			
			
			GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가

		}
		
		
			



;	MsgBox, 262144, Title, 스타일별로 정렬하기
	
	; 같은 스타일끼리 뭉쳐서 뽑기 위해서 정렬하기
	; 오래된 주문을 조금이라도 일찍 뽑기 위해 1열을 먼저 정렬한 뒤 색깔이 있는 4열(Columns)로 정렬한 뒤 최종적으로 스타일번호가 있는 3열로 정렬하기.
	xl.cells.sort(xl.columns(1), 1)
	xl.cells.sort(xl.columns(4), 1)
	xl.cells.sort(xl.columns(3), 1)
	GuiControl,,Progress, +5 ; 프로그래스 바 5 증가	




	Gui Destroy
	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 262144, Title, IT'S DONE


		
	

	Exitapp

	Esc::
	Exitapp	
	
	
	
	
	
	
	
	
	

!F2::
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible

		j = 1 ; 기준 스타일 번호의 줄 저장 변수
		k = 1 ; 비교할 스타일 번호의 줄 저장할 변수
		#ofStyle = 0 ; 같은 스타일이 몇개인지 세기 위해
		
		
		; 색깔이 있는 4열(Columns)로 정렬
		xl.cells.sort(xl.columns(4), 1)
		; 스타일 넘버가 있는 3열 정렬
		xl.cells.sort(xl.columns(3), 1)
		
		
		Loop{

			#ofRowToRead := "C" . j
			#ofRowToReadAddedOne :=  "C" . k
			
			
			standStyle# := Xl.Range(#ofRowToRead).Value
			Style#ToBeCompared := Xl.Range(#ofRowToReadAddedOne).Value
			
			;~ MsgBox, % standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k
			
			; 스타일 번호가 없는 빈칸이면 루프 끝내고 나가기
			if(standStyle# == ""){
				break
			}
						

			; standStyle# 변수와 Style#ToBeCompared 변수 값이 같으면
			; 다음 아이템 비교하기위해 루프 계속 진행하기
			IfEqual, standStyle#, %Style#ToBeCompared%
			{
;				MsgBox, % "아이템이 같음`n`n" . standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k

				; 다음줄로 넘어가기 위해
				k := k + 1
				
				; 같은 아이템이 몇개인지 확인하기 위해
				#ofStyle++ 
				
				continue
			}
			
			; 기준 스타일 번호와 지금 스타일 번호가 다르면 
			else IfNotEqual, standStyle#, %Style#ToBeCompared%
			{
				
;				MsgBox, % "아이템이 다름`n`n" . standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k
				
				; 10개 넘는 아이템만 빈줄 삽입해서 나누기
				;~ if(#ofStyle >= 2){ ; #################################################### 아이템이 1개든 2개든 스타일번호가 바뀔때마다 빈줄 넣고싶을때는 이것 사용하기 ####################################################
				if(#ofStyle >= 10){
					
;					MsgBox, % standStyle# . "`n" . Style#ToBeCompared . "`n" . "j : " . j . "`n" . "k : " . k
					
					; 10 개 넘는 아이템의 끝을 표시해주기 위해 엑셀에 빈줄 넣기
					Xl.Rows(k).EntireRow.Insert
					
					; 처음을 표시해주기 위해 빈줄 삽입해야하는데 만약 처음 표시할 빈줄 이전에 빈줄이 있다면 빈줄 삽입하지 않기. 만약 빈줄을 삽입하게 되면 빈줄이 연달아 2개가 되니까
					#ofRowToRead := "C" . j-1
					standStyle# := Xl.Range(#ofRowToRead).Value
					
					; 앞선 줄이 빈줄이 아닐때만 
					if(standStyle# != ""){
						
						; 처음을 다른 아이템과 나누기 위한 빈줄 삽입
						Xl.Rows(j).EntireRow.Insert
						
						; 10개 넘는 아이템의 처음과 끝을 표시해주기 위해 처음과 맨끝의 빈줄 2줄 입력했기 때문에 2 증가
						j := k + 2
						k := k + 2
						#ofStyle = 1						
;						MsgBox, % "j : " . j . "`n" . "k : " . k
						continue
					}
					
					; 10개 넘는 아이템이지만 첫줄은 삽입 않고 마지막 줄만 삽입했기 때문에 1 증가
					j := k + 1
					k := k + 1
					#ofStyle = 1
;					MsgBox, % "j : " . j . "`n" . "k : " . k
					continue
				}
				
				; 스타일 번호가 같은 갯수가 10개가 안되면 다음으로 넘어가기
				; k는 이미 스타일 번호가 다른 다음줄이기 때문에 기준 스타일 번호 줄을 k로 하기
				j := k
				#ofStyle = 1
;				MsgBox, % "다음의 j k 값`n`n" . "j : " . j . "`n" . "k : " . k
				continue
			}
			
			
			;~ GuiControl,,Progress, +4 ; 프로그래스 바 1씩 증가

		}
		





	SoundPlay, %A_WinDir%\Media\Ring06.wav
	MsgBox, 262144, Title, THE ITEMS HAVE BEEN DEVIDED.

	
	
	

	
	
	
	
	