#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include %A_ScriptDir%\lib\

#Include function.ahk
#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk



Loop
{



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
;	MsgBox % "last row: " XL_Last_Row(XL)  ;Last row
;	MsgBox, % i



	; 만약 i 값이 1이면 이전에 사용했던 파일이 안 닫히고 열려있는 것이므로 파일 닫고 프로그램 다시 시작하기
	IfEqual, i, 1
	{				
		; 파일이 끝났으니 메세지 띄우고 프로그램 다시 시작하기
		MsgBox, 262144, Old File Notification, IT'S A PROCESSED FILE`nPLEASE OPEN NEW BO LIST EXCEL FILE
		
		
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
		
		MsgBox, RefinedOrderID is : %RefinedOrderID%
		
		
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
		
		
		
	;	Start()
	;	Sleep 500

	;	MsgBox, 1
	;	continue

	}

}



























/*
while (Xl.Range("C" . A_Index).Value != "") {

	value := Xl.Range("C" . A_Index).Value
	RegExMatch(value, "imU)(\d*)\.", SubPat)
	CurrentRowNumber := A_Index
	;Xl.Range(Row . ":" . Row).Rows.Delete
	;XL_Row_Delete(XL,RG:="") ;can be out of order
	;COM_Invoke(pxl, "Rows[" m - minusRows++ "].Delete")
	;COM_Invoke(Xl, "Rows[].Delete")
	

;	Start()
;	Sleep 500

;	MsgBox, 1
;	continue

	MsgBox, %CurrentRowNumber%`n%SubPat1%
	Xl.Sheets(1).Range("A1").EntireRow.Delete	
	
}
*/

MsgBox, OUT






;i = 1
;XL_Row_Delete(XL,RG:=i) ;can be out of order



/*
; 백오더 숫자가 C9 부터 시작하기 때문에 i 에 8을 넣고 i + A_Index로 처리해서 값을 불러옴
; C Column 에 값이 없을때까지 반복
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; 값이 없어지면 루프를 끝내고 다시 시작하도록 reload 해줘야 됨
i = 8

while (Xl.Range("C" . i + A_Index).Value != "") {
		
	;while (Xl.Range("C" . 9).Value != "") {
	;Xl.Range("A" . A_Index).Value := value
	value := Xl.Range("C" . i + A_Index).Value
	RegExMatch(value, "imU)(\d*)\.", SubPat)
	CurrentRowNumber := i + A_Index

	Start()
	Sleep 500

	MsgBox, 1
	continue

	MsgBox, %CurrentRowNumber%`n%SubPat1%
}
*/














/*
DataA = 50449
StartRow = 0
EndRow = 0
RowCount = 0

Cell := xl.Range("C:C").Find(DataA)
StartRow := Cell.Row ;Find the starting row number
Loop
If Cell.Offset(A_Index+1, 0).Value <> DataA
{
	EndRow := Cell.Offset(A_Index, 0).Row
	break
}
Xl.Range("B" StartRow ":B" EndRow).Copy
return
*/





/*
File = %A_ScripDir%BO_ITEMS\D1190-1.xls ;Specify file
File = C:\Users\JODIFL4\Desktop\000000000\LAMBS\BO_ITEMS\D1190-1.xls ;Specify file
oWorkBook := COM_GetObject( File ) ;Create handle
FindThis = 50449 ;Value your looking for
;oRng := Com_Invoke(oWorkBook, "Range", "C1:D905") ;Range of cells to look inside
;COM_Invoke(oRng, "Find.Activate", FindThis) ;Search for value inside this range of cells
;[color=red]COM_Invoke(oWorkBook,"Range[C1].Find[FindThis].Activate")[/color]
;COM_Invoke(oWorkBook,"[color=red]ActiveSheet[/color].Range[D2].Find[" FindThis "].Activate")
;COM_Invoke(oWorkBook,"[color=red]ActiveSheet[/color].Range[D2].Find[" FindThis "].Activate")
*/


/*

; 숫자가 C9 부터 시작하기 때문에 i 에 8을 넣고 i + A_Index로 처리해서 값을 불러옴
i = 8
MsgBox, % i
while (Xl.Range("C" . i + A_Index).Value != "") {
;while (Xl.Range("C" . 9).Value != "") {
;Xl.Range("A" . A_Index).Value := value
value := Xl.Range("C" . i + A_Index).Value
RegExMatch(value, "imU)(\d*)\.", SubPat)
CurrentRowNumber := i + A_Index

Start()
Sleep 500

MsgBox, 1
continue

MsgBox, %CurrentRowNumber%`n%SubPat1%
}

*/

ExitApp

GuiClose:
ExitApp
  
Esc::
 Exitapp
 return

