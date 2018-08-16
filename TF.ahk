; ############################################################################################################
; ### 엑셀에서 오더 정보 읽어서 N41에서 Sales Order 만든 후 eLAMBS 통해서 정보 다 닫은 후 (FG나 LAS 업데이트 하기)
; ############################################################################################################

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)


#Include %A_ScriptDir%\lib\

#Include EXCEL.ahk
#Include N41.ahk
#Include eLAMBS.ahk
;~ #Include CommWeb.ahk


; 아이템이 여러개인 것 표시하기 위해
global preOrderId


; 엑셀에서 읽은 값 저장하는 배열 선언
Str_Total_ExcelInfo := object()



E_driver := new EXCEL
N_driver := new N41
eL_driver := new eLAMBS




Loop{
	
	N41_login_wintitle := "ahk_class FNWND3126"
	IfWinNotExist, %N41_login_wintitle%
	{
		MsgBox, 262144, Alert, THERE IS NO N41 WINDOW`n`nIF OK BUTTON CLICKED, THE APPLICATION WILL BE RELOADED
		Reload
	}


	; 엑셀에서 정보 읽기. 같은 고객의 모든 정보가 다중배열로 반환됨
	Str_Total_ExcelInfo := E_driver.GetInfoFromExcelThenPutThatInAArrayForTF()

;~ /*
	; eLAMBS 열어서 값 얻어오기
	; 다중 배열인 Str_Total_ExcelInfo 안에 값이 몇 개가 들었든 가장 처음 배열의 처음 값은 Order Id니까 Order Id 넘기면서 메소드 호출하기
	;~ ShippingAddrOneLAMBS := eL_driver.Get_SOInfo_ofLAMBS(Str_Total_ExcelInfo[1][1])
	Str_Addr := eL_driver.Get_SOInfo_ofLAMBS(Str_Total_ExcelInfo[1][1])

	;~ MsgBox, % "ShippingAddrOneLAMBS : " . ShippingAddrOneLAMBS


	; Str_Total_ExcelInfo 의 각 끝에(8번째) eLAMBS에서 읽어온 주소값 추가하기
	Loop % Str_Total_ExcelInfo.Maxindex(){
		;~ Str_Total_ExcelInfo[A_Index].Insert(ShippingAddrOneLAMBS)
		Str_Total_ExcelInfo[A_Index].Insert(Str_Addr[1]) ; Billing Addr
		Str_Total_ExcelInfo[A_Index].Insert(Str_Addr[2]) ; Shipping Addr
	}
*/




	
	; 배열에 들어있는 스타일 갯수만큼만 루프 돌려서 n41에서 Sales Order 만들기. 스타일이 2개면 2번만 돌리기
	Loop % Str_Total_ExcelInfo.Maxindex(){
		N_driver.MakeSalesOrderUsingInfoFromLAMBS(Str_Total_ExcelInfo[A_Index])
	}
	
	


	; 배열에 있는 아이템들의 체크박스에 체크하기
	; 위에서 값을 얻은 후 열려있는 eLAMBS를 이용.
	Loop % Str_Total_ExcelInfo.Maxindex(){
		eL_driver.CheckCheckBoxOfItems(Str_Total_ExcelInfo[A_Index])	
	}



	; eLAMBS 마무리하기
	; 아이템 상태를 void로 바꾸고 메모창에 TRANSFERRED TO N41 적은 후 저장하고 끝내기
	eL_driver.WrapUpeLAMBS()
	
	
	;~ SoundPlay, %A_WinDir%\Media\Ring06.wav
	;~ MsgBox, PROCESSING FINISH`n`nMOVE TO NEXT CUSTOMER IF OK BUTTON CLICKED

	;~ WinClose, eLAMBS | Edit Sales Order - Google Chrome
	
;MsgBox, FINISH

}











/*
i = 1
j = 1
;~ /* 배열로부터 읽기 첫 번째 방법
Loop % Str_Total_ExcelInfo.Maxindex(){
	Loop, 8{ ; 쓸 수 있는 유효값이 8개니까
		MsgBox % "Element number " . A_Index . " is " . Str_Total_ExcelInfo[i][j]
		j++
	}
	
	i++
	j = 1
}
*/





ExitApp





















F6::
;~ /*
Str_Total_ExcelInfo := E_driver.GetInfoFromExcelThenPutThatInAArrayForTF()

; 위에서 값을 얻은 후 열려있는 eLAMBS를 이용.
; 배열에 있는 아이템들의 체크박스에 체크하기
Loop % Str_Total_ExcelInfo.Maxindex(){	
	eL_driver.CheckCheckBoxOfItems(Str_Total_ExcelInfo[A_Index])
}
*/
	; 아이템 상태를 void로 바꾸기, 메모창에 TRANSFERRED TO N41 적은 후 저장하기
	eL_driver.WrapUpeLAMBS()
	
return








F7::
WinTitle := "ahk_class FNWND3126"

;~ ControlClick [, Control-or-Pos, WinTitle, WinText, WhichButton, ClickCount, Options, ExcludeTitle, ExcludeText]
ControlClick, Edit66, %WinTitle%
MsgBox
;~ ControlGetText, OutputVar [, Control, WinTitle, WinText, ExcludeTitle, ExcludeText]
ControlGetText, OutputVar, Edit66, %WinTitle%
MsgBox, % OutputVar
;~ ControlSetText [, Control, NewText, WinTitle, WinText, ExcludeTitle, ExcludeText]
ControlSetText, Edit66, abcdef, %WinTitle%

MsgBox

return









Esc::
ExitApp