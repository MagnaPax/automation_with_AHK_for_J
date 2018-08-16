#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

#Include EXCEL.ahk

global Arr_Excel



;~ Test := returnMultipleArrays()
;~ MsgBox, % Test[1][1] "`, " Test[1][2] . "`n" . Test[2][1] "`, " Test[2][2] . "`n" . Test[3][1] "`, " Test[3][2]




;~ Arr_Excel := object()



Array := GetInfoFromExcel2()

MsgBox, GetInfoFromExcel2 out

MsgBox, % Array[1]

	
	Loop{
		
		if( Arr_Excel%A_Index% == "")
			break
		
		MsgBox, % Arr_Excel%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index% . "`n" . Arr_Excel%A_Index%_%A_Index%_%A_Index%_%A_Index%

	}
	

MsgBox, PAUSE


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
		


	}





GetInfoFromExcel()


Exitapp	

Esc::
 Exitapp