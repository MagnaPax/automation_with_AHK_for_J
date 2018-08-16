#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

#Include FindTextFunctionONLY.ahk
#Include FG.ahk

#Include LAMBS.ahk
#Include CommonLAMBSProcessing.ahk
#Include N41.ahk
#Include CommonN41Processing.ahk


global #ofCC_counter



L_driver := new LAMBS
N_driver := new N41
F_driver := New FG

Arr_CSOS := object()
Arr_CC := object()
Arr_FGInfo := object()


/*
; [##1##] CSOS 에서 정보 얻기
;~ Arr_CSOS := L_driver.getInfoFromCSOS()

; [##2##] 램스에서 카드 정보 읽어서 배열에 저장하기
;~ Arr_CC := L_driver.ReadingCCInfoFromLAMBS()

;~ MsgBox, % "value of variable #ofCC_counter : " #ofCC_counter
;~ MsgBox, % Arr_CC[1][3]
;~ MsgBox, % Arr_CC[1][5]


; [##3##] 램스에서 읽은 카드정보 N41에 입력하기
;~ IsItFromFG = 1
;~ ReadCCInfo_then_PutThatinN41CCWindow(Arr_CC, Arr_CSOS, N_driver, IsItFromFG)
*/




Arr_CSOS[6] := "MTR1E199323C8" ; 빌링과 쉬핑 주소 다른 것
;~ Arr_CSOS[6] := "MTR1DFEE64CAB" ; 주소에 add2 있는 것

; [##4##] FG 오더면 FG 에서 주문 페이지에서 카드 정보 읽은 후 N41에 넣기
if(RegExMatch(Arr_CSOS[6], "imU)MTR")){
	
	; FG에서 카드 정보 읽어서 배열에 저장하기
	Arr_FGInfo := F_driver.GettingInfoFromCurrentPage(Arr_CSOS[6])
	
	
	; 여러 값이 들어있는 Arr_FGInfo 배열에서 CC 정보가 들어있는 Arr_FGInfo[1](첫 번째 배열값) Arr_CC 배열로 복사
	;~ Arr_CC := Arr_FGInfo[1].Clone()
	Arr_CC := Arr_FGInfo.Clone()
	
	MsgBox, % "It's out of all methods`n`n`n`n" . "Arr_CC[1][1] : " . Arr_CC[1][1] . "`nArr_CC[1][2] : " . Arr_CC[1][2] . "`nArr_CC[1][3] : " . Arr_CC[1][3] . "`nArr_CC[1][4] : " . Arr_CC[1][4] . "`nArr_CC[1][5] " . Arr_CC[1][5] . "`nArr_CC[1][6] " . Arr_CC[1][6] . "`nArr_CC[1][7] " . Arr_CC[1][7] . "`nArr_CC[1][8] " . Arr_CC[1][8] . "`nArr_CC[1][9] " . Arr_CC[1][9] . "`nArr_CC[1][10] " . Arr_CC[1][10] . "`nArr_CC[1][11] " . Arr_CC[1][11]

	;~ MsgBox, % "Arr_CC[1][5] : " . Arr_CC[1][5] . "`n`n" . "Arr_CC[1][6] : " . Arr_CC[1][6]
	
	
	; FG에서 읽은 카드정보 N41에 입력하기
	IsItFromFG = 1
	ReadCCInfo_then_PutThatinN41CCWindow(Arr_CC, Arr_CSOS, N_driver, IsItFromFG)
}












ReadCCInfo_then_PutThatinN41CCWindow(Arr_CC, Arr_CSOS, N_driver, IsItFromFG){	
	
	MsgBox, ReadCCInfo_then_PutThatinN41CCWindow Method in
	MsgBox, % "Arr_CC[1][1] : " . Arr_CC[1][1] . "`nArr_CC[1][2] : " . Arr_CC[1][2] . "`nArr_CC[1][3] : " . Arr_CC[1][3] . "`nArr_CC[1][4] : " . Arr_CC[1][4] . "`nArr_CC[1][5] " . Arr_CC[1][5]
	
		
	i = 0 ; 읽어들인 카드 갯수가 몇 개인지 세기 위해. i값의 갯수만큼 N41에 저장한다
	j = 1 ; 카드 번호 카운터. j값이 1이면 첫 번째 카드 정보 2면 두 번째 카드 정보

	;~ loop, 10{ ; 신용카드 갯수
	;~ Loop, % Array.Maxindex(){
	Loop{
		
		; 만약 카드 정보가 없으면 루프 탈출
		if(Arr_CC[j][3] == ""){
			MsgBox, No CC info
			break
		}	
		
		; 11번째 United States 값 다음인 12번째에 Shipping ADD 의 전화번호 넣기
		Arr_CC[j][12] := Arr_CSOS[5]
		

		; 이전 카드 번호와 같은 카드 번호가 들어있으면 중복된 정보가 들어있다는 뜻이므로 루프 중단
		if(Arr_CC[j][3] == previousCCNum){
			break
		}

	/*
		; N41 열어서 저장하기
		Loop, 12{ ; 카드 한 개에 들어있는 카드 정보 갯수는 11개니까. 11번째 값은 United States 이거나 정보가 들어있지 않거나 대부분 둘 중 하나. 아직 해외 발급 카드는 못 본듯
			MsgBox % "Element number " . A_Index . " is " . Arr_CC[j][A_Index]
			;~ N_driver.PutInfoInN41(Arr_CC[j])
		}
	*/

		; N41 에 카드 정보 입력하기
		N_driver.PutInfoInN41(Arr_CC[j])
		
		
		; 중복된 카드 체크하기 위해 
		previousCCNum := Arr_CC[j][3]
		
		
		j++
		i++ ; 읽어들인 카드 갯수가 몇 개인지 세기 위해. i값의 갯수만큼 N41에 저장할 것이다
		
		if(IsItFromFG == 1)
			break
	}

	MsgBox, % "A number of CC of this customer : " i	
	
}










Exitapp

Esc::
Exitapp