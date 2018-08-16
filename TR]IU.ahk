; ############################################################################################################
; 아이템 자동으로 업데이트 하기 위해 Invoice && C.M Detail Register 엑셀 파일 전처리 하기

; 엑셀 파일에 송장 번호(Tracking #)가 있는 것들이 자동으로 처리된 것들.(= 프로그램이 아이템을 업데이트 해야되는 인보이스들)
; N41이 자동으로 처리한 후 송장번호를 임의로 생성한 것들은 엑셀에 송장번호가 있기 때문.
; 사람이 수동으로 처리한 것은 인보이스의 메모란에 송장번호를 넣지 기계에 직접 넣을 수가 없다
; ############################################################################################################














#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.




#Include %A_ScriptDir%\lib\

#Include EXCEL.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk

global Progress



E_driver := new EXCEL


; 아이템 업데이트 위해 배송정보 들어있는 Invoice && C.M Detail Register 파일에서 빈줄 삭제하고 쓸데없는 열(column) 지우기
;~ E_driver.DeleteEmptyColumnToUpdateItems()
DeleteEmptyColumnToUpdateItems()


; void 랑 cancel 메모가 들어있는 줄만 남기고 지우기
;~ E_driver.DeleteColumnExceptOnesHavingVoidMemo()







	

	; 아이템 업데이트 위해 배송정보 들어있는 Invoice && C.M Detail Register 파일에서 빈줄 삭제하고 쓸데없는 열(column) 지우기
	DeleteEmptyColumnToUpdateItems(){

		; 엑셀 창이 열려있지 않으면 열릴때까지 경고창 표시하기
		CommEXCEL.CheckingPresenceOfExcelFile()


		; 열려있는 엑셀 창 사용하기
		Xl := ComObjActive("Excel.Application")
		Xl.Visible := True ;by default excel sheets are invisible
		
		
		
		
		; 아이템 찾는 동안 보여줄 프로그래스 바 
		TotalLoops = 3000
		Gui, -Caption +AlwaysOnTop +LastFound
		Gui, Add, Text, x12 y9 w170 h20 , P  R  O  C  E  S  S  I  N  G  .  .  .
		Gui, Add, Progress, w410 Range0-%TotalLoops% vProgress
		Gui, Show, w437 h84, SEARCHING ITEMS

		
		
		
		
		
	
		i = 1
		loop{
			
			CellNumberCurrentlyBeingBeenHandled := "A" . i
			
			; A1 셀에 있는 값 IsItEmptyOrNot 변수에 저장하기
			;~ IsItEmptyOrNot := Xl.Range("A1").Value
			IsItEmptyOrNot := Xl.Range(CellNumberCurrentlyBeingBeenHandled).Value
			
			
			
			
			; 맨 마지막 A1 셀에는 'Grand Total' 값이 있기 때문에 그것이 나오면 여러가지 처리한 뒤 메소드 끝내기
			if IsItEmptyOrNot contains Grand
			{			
MsgBox, % "`n`n`n`n`n`n`n`n`n`n`n" . 1
				
				; 맨 마지막 두 줄 삭제하기
				; Grand Total 있는 맨 마지막 전 줄과
				; *Season from  Style Master* 쓰여져 있는 제일 마지막 줄 지우기
				Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				
				;~ CellNumberCurrentlyBeingBeenHandled := "A" . ++i
				;~ Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				;~ GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
MsgBox, % "`n`n`n`n`n`n`n`n`n`n`n" . 2
				
				; 맨 마지막 이전 줄 삭제하기
				;~ CellNumberCurrentlyBeingBeenHandled := "A" . --i
				CellNumberCurrentlyBeingBeenHandled := "A" . ++i
				Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
MsgBox, % "`n`n`n`n`n`n`n`n`n`n`n" . 3
				
				
				; 맨 처음에 남은 제목 줄 삭제하기
				Xl.Sheets(1).Range("A1").EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				

				; 쓸 데 없는 행 삭제하기
				;~ XL_Col_Delete(XL,RG:="D:F|H|K|N:R") ;Delete columns
				;~ XL_Col_Delete(XL,RG:="C|E:N|P:AI|AK:AX|AZ:BX|BZ:CC") ;Delete columns
				
				
				;~ XL_Col_Delete(XL,RG:="C|E:N|P:Q|R:AA|AC:AD") ;Delete columns
				;~ XL_Col_Delete(XL,RG:="C|E:N|P:Q") ;Delete columns
				XL_Col_Delete(XL,RG:="C|E:N|P:Q|S:Y|AA|AC:AD") ;Delete columns
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				
			

				
				; Tracking # 열 넓이 설정하기
				;~ XL_Col_Width_Set(XL,RG:="A=10|C=8|D=30|F=15|G=20")
				;~ XL_Col_Width_Set(XL,RG:="A=10|B=10|C=10|D=30|E=20|F=20")
				XL_Col_Width_Set(XL,RG:="A:C=10|D=30|E=5|F=20|G=30")
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	



				; Tracking # 크기대로 정렬하기
				; 7 열(Columns)을 정렬하기
				;~ xl.cells.sort(xl.columns(7), 1)
				xl.cells.sort(xl.columns(7))
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
				

				MsgBox % "`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n" . "last row: " XL_Last_Row(XL)  ;Last row
				
				
				; 맨 마지막 줄 삭제하기
				; *Season from  Style Master* 쓰여져 있는 마지막 줄 지우기
				
				Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가					
				


				; 엑셀 파일에 송장 번호(Tracking #)가 있는 것들이 자동으로 처리된 것들.(= 프로그램이 아이템을 업데이트 해야되는 인보이스들)
				; N41이 자동으로 처리한 후 송장번호를 임의로 생성한 것들은 엑셀에 송장번호가 있기 때문.
				; 사람이 수동으로 처리한 것은 송장번호를 넣지 기계에 직접 넣을 수가 없다. 대신 메모란에 적는다
				; Tracking # 값이 들어있는 행 루프 돌아가며 모두 지우기
				
				
				; ########################################################################################################################################################################################################################################
				; 아래 loop 바꿔야 됨
				; 지금은 TRACKING # 있는것을 지우는데 반대로 TRACKING # 가 없는 것들을 지워야 됨
				; ########################################################################################################################################################################################################################################				
				Loop{
					
					; G 열의 첫 번째 행의 값, 즉 첫번째 Tracking # 값을 var1 변수에 넣기					
					var1 := Xl.Range("G1").Value
					
					; Tracking # 가 있다면 지우기
					if(var1){
						Xl.Sheets(1).Range("G1").EntireRow.Delete
						GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
					}
					; 다 지웠으면 루프 빠져나오기
					else
						break
				}
				
				
				; CUSTOER PO 로 정렬하기
				; 6 열(Columns)을 정렬하기
				xl.cells.sort(xl.columns(6), 1)
				
				

	/*			; 원하는 폴더에 원하는 이름으로 저장하기
				; https://autohotkey.com/board/topic/94976-save-as-in-excel/
				book_name := Xl.ActiveWorkbook.name
				SplitPath, book_name,,, ext
				;~ Xl.ActiveWorkbook.SaveAs(directory "\" new_file_name "." ext)
				Xl.ActiveWorkbook.SaveAs(%A_ScriptDir%\CreatedFiles "\" new_file_name "." xls)			
				
	*/

				Gui Destroy
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, IT'S DONE				
						
						
				return
			}
			
			

			; A1에 값이 없으면 그 줄 삭제하기
			if(IsItEmptyOrNot == ""){
				;~ Xl.Sheets(1).Range("A1").EntireRow.Delete
				Xl.Sheets(1).Range(CellNumberCurrentlyBeingBeenHandled).EntireRow.Delete
				GuiControl,,Progress, +2 ; 프로그래스 바 1씩 증가	
			}
			else
				i++



		}

		return
	}
	








;~ XL_Last_Row(PXL){
 ;~ Return, PXL.Application.ActiveSheet.UsedRange.Rows(PXL.Application.ActiveSheet.UsedRange.Rows.Count).Row
;~ }













ExitApp

Esc::
ExitApp