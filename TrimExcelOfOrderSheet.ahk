#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

;~ #Include function.ahk
#Include [Excel]_InsertingORDeletingAndSettingHeightOFRowsINExcel.ahk
#Include [Excel]_ObtainFirstrow_Lastrow_#UsedrowsfromExcel.ahk
#Include [Excel]_Joe Glines'sExcelFunctions.ahk



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
	

	; 열려있는 엑셀 창 사용하기
	Xl := ComObjActive("Excel.Application")
	Xl.Visible := True ;by default excel sheets are invisible
	
	; 1~8줄(Row) 지우기
	Xl.Sheets(1).Range("A1:A8").EntireRow.Delete
	
	; A~B, D, F, H~X 열(Columns) 지우기
	XL_Col_Delete(XL,RG:="A:B|D|F|H:X") ;Delete columns	
	
	; A~C 열의 글자 크기 폰트 설정하기
	;~ XL_Format_Font(XL,RG:="A:C",Font:="Book Antiqua",Size:=11) ;Arial, Arial Narrow, Calibri,Book Antiqua
	XL_Format_Font(XL,RG:="A:C",Font:="Arial",Size:=11) ;Arial, Arial Narrow, Calibri,Book Antiqua
	
	
	; 줄의 높이 설정하기
	;~ XL_Row_Height(XL,RG:="1:4=-1|10:13=50|21=15") ; 1~4 줄의 높이는 -1, 10~13의 높이는 50, 21의 높이는 15
	XL_Row_Height_(XL,RG:="1:50=17")


	; 열의 넓이 설정하기
	;~ XL_Col_Width_Set(XL,RG:="A:B=-1|D:F=-1|H=15|K=3") ;A~B 는 -1, D~F는 -1, K는 3 ;-1 is auto
	XL_Col_Width_Set(XL,RG:="A=10|B=25|C=70")
	
	
	; 셀 안의 글자 위치 정열 
	;~ XL_Format_HAlign(XL,RG:="A1:C1",h:=1) ;1=Left 2=Center 3=Right	
	;~ XL_Format_VAlign(XL,RG:="A1:C1") ;1=Top 2=Center 3=Distrib 4=Bottom	
	XL_Format_HAlign(XL,RG:="A:B",h:=2) ;1=Left 2=Center 3=Right
	XL_Format_VAlign(XL,RG:="A:B",v:=2) ;1=Top 2=Center 3=Distrib 4=Bottom
	
	
	; 인쇄하기
;	Xl.ActiveSheet.PrintOut ; 미리보기 없이 현재 엑셀 화면 곧바로 인쇄하기
	;~ COM_Invoke(VAR_PWB, "document.parentWindow.print")

	; 여러대 프린터로 인쇄하기
	;~ Xl.ActiveSheet.PrintOut(From := 1, To := 1, Copies := 2, Preview := m, ActivePrinter := "\\sbs2k3\Xerox WC7300v1.0SP1 (EFI)", PrintToFile := m, Collate := m, PrToFileName := m, IgnorePrintAreas := m) ; 첫번째 프린터
	;~ Xl.ActiveSheet.PrintOut(From := 1, To := 1, Copies := 1, Preview := m, ActivePrinter := "PDF24 PDF", PrintToFile := m, Collate := m, PrToFileName := m, IgnorePrintAreas := m) ; 두 번째 프린터

	
	; 아래 코드는 그냥 예제로 넣은것 
	; 1 열(Columns)을 정렬하기
;	xl.cells.sort(xl.columns(1), 1)
	


;***********************Column Delete********************************.
XL_Col_Delete(PXL,RG=""){
	for j,k in StrSplit(rg,"|")
		(instr(k,":")=1)?list.=k ",":(list.=k ":" k ",") ;need to make for two if only 1 col
	PXL.Application.ActiveSheet.Range(SubStr(list,1,(StrLen(list)-1))).Delete ;use list but remove final comma
}

;***********************set size, type, ********************************.
XL_Format_Font(PXL,RG="",Font="Arial",Size="11"){
	PXL.Application.ActiveSheet.Range(RG).Font.Name:=Font
	PXL.Application.ActiveSheet.Range(RG).Font.Size:=Size
}

 
;***********************Row Height********************************.
XL_Row_Height_(PXL,RG=""){
	for k, v in StrSplit(rg,"|") ;Iterate over array
		(StrSplit(v,"=").2="-1")?(PXL.Application.ActiveSheet.rows(StrSplit(v,"=").1).AutoFit):(PXL.Application.ActiveSheet.rows(StrSplit(v,"=").1).RowHeight:=StrSplit(v,"=").2)
}

;***********column width*******************
XL_Col_Width_Set(PXL,RG=""){
	for k, v in StrSplit(rg,"|") ;Iterate over array
		(StrSplit(v,"=").2="-1")?(PXL.Application.ActiveSheet.Columns(StrSplit(v,"=").1).AutoFit):(PXL.Application.ActiveSheet.Columns(StrSplit(v,"=").1).ColumnWidth:=StrSplit(v,"=").2)
}

XL_Format_HAlign(PXL,RG="",h="1"){ ;defaults are Right bottom
	IfEqual,h,1,Return,PXL.Application.ActiveSheet.Range(RG).HorizontalAlignment:=-4131 ;Left
	IfEqual,h,2,Return,PXL.Application.ActiveSheet.Range(RG).HorizontalAlignment:=-4108 ;Center
	IfEqual,h,3,Return,PXL.Application.ActiveSheet.Range(RG).HorizontalAlignment:=-4152 ;Right
}

; Excel cell alignment with AutoHotkey
XL_Format_VAlign(PXL,RG="",v="1"){		
	IfEqual,v,1,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4160 ;Top
	IfEqual,v,2,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4108 ;Center
	IfEqual,v,3,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4117 ;Distributed
	IfEqual,v,4,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4107 ;Bottom
}



Exitapp

Esc::
Exitapp