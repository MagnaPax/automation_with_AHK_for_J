#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;~ #Include function.ahk



	; 마우스 커서는 사용자의 물리적 마우스 이동에 반응하지 않습니다
	BlockInput, MouseMove
	BlockInput, MouseMoveOff
	
	
	CoordMode, mouse, relative
	
	;LAMBS활성화 후 화면 초기화
	Start()
	OpenCreateSalesOrdersSmallTab()
	Sleep 1000

	; LAMBS에 있는 인보이스 밸런스 값 읽어오기
;	GetInvoiceBalanceOnLAMBS()


	;CC 버튼 찾고 클릭
	FindCCButtonAndClickIt()
	WinWaitActive, Credit Card ( P999131 ) ; CC 창 뜰 때까지 기다리기
	
	
	i = 1 ; i 값이 증가한 만큼 배열 갯수가 만들어 졌다
	loop{
		
		Array%i% := getCCInfoFromCCWindowOfLAMBS() ; LAMBS의 cc 창에서 정보 얻은 후 배열에 저장
		
		; 카드 정보가 들어있는 배열의 세번째 값에 아무것도 들어있지 않으면 LAMBS예는 CC값이 없다는 뜻이고 루프를 빠져나간다
		if(Array%i%[3] == ""){
			MsgBox, No CC info in LAMBS
			WinClose, Credit Card ( P999131 )
			break
		}
		
		; 이전 카드 번호와 같은 카드 번호가 들어있으면 중복된 정보가 들어있다는 뜻이므로 중복된 마지막 배열을 삭제 후 루프 중단
		if(Array%i%[3] == previousCCNum){
			Array%i%.remove() ; previousCCNum 와 같은 값이 들어있는 배열은 중복된 것이니 지워주기
			--i ; if 들어오기 전에 값을 한 번 읽어줬기 때문에 배열의 갯수를 카운트 하는 i 값을 줄이기
			WinClose, Credit Card ( P999131 )
			break
		}
	
		; CC 창 기다렸다가 아무데나 클릭해서 한 칸 내려서 다음 CC 정보 읽도록 처리해주기
		ToMoveNextCCInfo()
		
		previousCCNum := Array%i%[3]
		
		i++ ; 배열이 몇 개 만들어 졌는지 세기 위해
	}
	
	MsgBox % "i : " i
	
	; 한 고객이 10개의 카드를 갖고 있을리는 없으니까 넉넉하게 10개 리턴하기
	Array := [Array1, Array2, Array3, Array4, Array5, Array6, Array7, Array8, Array9, Array10]
	
i = 0 ; 읽어들인 카드 갯수가 몇 개인지 세기 위해. i값의 갯수만큼 N41에 저장한다
j = 1 ; 카드 번호 카운터. j값이 1이면 첫 번째 카드 정보 2면 두 번째 카드 정보
loop, 10{ ; 신용카드 갯수

	; 이전 카드 번호와 같은 카드 번호가 들어있으면 중복된 정보가 들어있다는 뜻이므로 루프 중단
	if(Array[j][3] == previousCCNum){
		break
	}
	
	; N41 열어서 저장하기
	Loop, 11{ ; 카드 한 개에 들어있는 카드 정보 갯수는 11개니까. 11번째 값은 United States 이거나 정보가 들어있지 않거나 대부분 둘 중 하나. 아직 해외 발급 카드는 못 본듯
		MsgBox % "Element number " . A_Index . " is " . Array[j][A_Index]
		;~ PutInfoInN41(Array[j][A_Index])
	}
	
	; 중복된 카드 체크하기 위해 
	previousCCNum := Array[j][3]
	
	j++
	i++ ; 읽어들인 카드 갯수가 몇 개인지 세기 위해. i값의 갯수만큼 N41에 저장할 것이다
}

MsgBox, % "A number of CC of this customer : " i

/*
	loop{
		
		; PO 번호 CurrentPONumber 변수에 저장하기	
		DllCall("SetCursorPos", int, 994-8, int, 378-8)  ; 첫 번째 숫자는 X-좌표이고 두 번째 좌표는 Y입니다 (화면에 상대적입니다).
	;	MouseMove, 994, 378
		Sleep 100
		MouseGetPos, , , , control, 1
		ControlGetText, CurrentPONumber, %control%, %wintitle%		
		
		Array%i% := getCCInfoFromFG()
	}
*/	
	
	
	
	Loop, %i%{ ; i 값과 만들어진 배열 갯수가 같기 때문에 i값 만큼만(배열 갯수만큼만) 루프 돌린다
;		Array%A_Index%
	}
	
	MsgBox, method out
	
	
	
	
	
	Loop, %i%{ ; i 값과 만들어진 배열 갯수가 같기 때문에 i값 만큼만(배열 갯수만큼만) 루프 돌린다
		PutInfoInN41(Array%A_Index%)
	}
	
	MsgBox, method out



	j = 1 ; Array1 부터 시작해서 그 다음에 Array2 이런 식으로 배열 번호를 늘리면서 값을 읽기 위해	
	Loop, %i%{ ; i 값과 만들어진 배열 갯수가 같기 때문에 i값 만큼만(배열 갯수만큼만) 루프 돌린다
		Loop % Array%j%.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Array%j%[A_Index]
		}
		j++
	}
	
	MsgBox, loop out





	PutInfoInN41(Array){
		
		WinActivate, ahk_class FNWND3126
		WinClose, Credit Card Management
		
		
		; 카드 아이콘 클릭
		Text:="|<CCIcon>*147$14.zzzzwDz000000001000000000000U"
		if ok:=FindText(697,133,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		}
		
		
		WinWaitActive, Credit Card Management
		Sleep, 500
		
		
		; 카드 정보 추가 하기 위해 New 버튼 클릭
		Text:="|<New Button>*129$37.3k00001800000Y0MU00G0AE03vw59mH7y2Z5Nbz1+yezz0bER0w0Fg6kS08Hm8D00000700000U"
		if ok:=FindText(578,284,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  
		  Sleep 500
		}


		
		Loop % Array.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
		}
		
		return
	}













	
	; CC 창 기다렸다가 아무데나 클릭해서 한 칸 내려서 다음 CC 정보 읽도록 처리해주기
	ToMoveNextCCInfo(){
		
		WinActivate, Credit Card ( P999131 ) ; 아래 DllCall 이 화면에 상대적이기 때문에 활성화 해주기

		; 아무데나 클릭해서 화살표 내려주면 다음 cc로 넘어가기 때문에 
		Text:="|<Default>*152$49.yDrkV2EzlY20MV828G10IEY1490U88G0W4yT0Y90F2E84G4U8V843x2E4FY212V82DXx11D7l4"

		if ok:=FindText(3020,595,150000,150000,0,0,Text)
		{
		  CoordMode, Mouse
		  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
		  MouseMove, X+W//2, Y+H//2
		  Click
		  Send, {Down}
		}

		Sleep 1000
			
		return
	}


	; 램스의 CC 창에서 CC 정보 읽기
	getCCInfoFromCCWindowOfLAMBS(){
		
		windowtitle := "Credit Card ( P999131 )"
		
		; CC Type 저장
		ControlGetText, CCType, WindowsForms10.EDIT.app.0.378734a17, %windowtitle%
		
		; Name on Card 저장
		ControlGetText, CCName, WindowsForms10.EDIT.app.0.378734a13, %windowtitle%
		StringUpper, CCName, CCName ; 대문자로 바꾸기
		
		; CC번호 저장
		ControlGetText, CCNumbers, WindowsForms10.EDIT.app.0.378734a12, %windowtitle%

		; CVS 저장
		ControlGetText, CVV, WindowsForms10.EDIT.app.0.378734a11, %windowtitle%
		
		; 만료일 저장
		ControlGetText, ExpDate, WindowsForms10.EDIT.app.0.378734a3, %windowtitle%
		
		ExpDate := Refine_ExpDate(ExpDate)

		; ADD 1 저장
		ControlGetText, ADD1, WindowsForms10.EDIT.app.0.378734a9, %windowtitle%
		StringUpper, ADD1, ADD1 ; 대문자로 바꾸기

		; ADD 2 저장
		ControlGetText, ADD2, WindowsForms10.EDIT.app.0.378734a8, %windowtitle%
		StringUpper, ADD2, ADD2 ; 대문자로 바꾸기

		; CITY 저장
		ControlGetText, CITY, WindowsForms10.EDIT.app.0.378734a7, %windowtitle%

		; STATE 저장
		ControlGetText, STATE, WindowsForms10.EDIT.app.0.378734a6, %windowtitle%

		; ZIP 저장
		ControlGetText, ZIP, WindowsForms10.EDIT.app.0.378734a5, %windowtitle%

		; COUNTRY 저장
		ControlGetText, COUNTRY, WindowsForms10.EDIT.app.0.378734a4, %windowtitle%


;		MsgBox, % CCType . "`n`n" . CCName . "`n`n" . CCNumbers . "`n`n" . CVV . "`n`n" . ExpDate . "`n`n" . ADD1 . "`n`n" . ADD2 . "`n`n" . CITY . "`n`n" . STATE . "`n`n" . ZIP . "`n`n" . COUNTRY
		
		Array := [CCType, CCName, CCNumbers, CVV, ExpDate, ADD1, ADD2, CITY, STATE, ZIP, COUNTRY]
		Sleep 1000
		
		return Array
		
	}















	; 만료일 4자리 숫자로 만들기
	Refine_ExpDate(ExpDate){
		
		;~ ExpDate = 8/2022
		;~ ExpDate = 10/2022
		;~ ExpDate = 04 / 2020
		;~ ExpDate = 12 / 2020


		StringReplace, ExpDate, ExpDate, %A_SPACE%, , All ; 모든 스페이스 제거
		StringReplace, ExpDate, ExpDate, /, , All ; 모든 / 제거
		ExpDate := Trim(ExpDate)
		if(StrLen(ExpDate) == 5)
			ExpDate := "0"ExpDate
		
		StringLeft, leftOf, ExpDate, 2
		StringRight, RightOf, ExpDate, 2
		
		ExpDate := % leftOf . RightOf
		
		
;		MsgBox, % ExpDate
		
		
		return ExpDate
	}






	;LAMBS활성화 후 화면 초기화 하기
	Start(){
		
		;LAMBS Window 활성화 하기
		WinActivate, LAMBS -  Garment Manufacturer & Wholesale Software
		windowtitle = LAMBS -  Garment Manufacturer & Wholesale Software
		CheckTheWindowPresentAndActiveIt(windowtitle)

		;Hide All 클릭해서 메뉴 바 없애기
		ClickAtThePoint(213, 65)
		
		return
	}
	
/*
; LAMBS에 있는 인보이스 밸런스 값 읽어오기
GetInvoiceBalanceOnLAMBS()
{

	;LAMBS활성화 후 화면 초기화
	Start()
	
	;Invoice Balance 값 얻기
	MouseClick, l, 870, 594
	Send, ^a^c
	Sleep 700
	InvoiceBalance := Clipboard
	
;	MsgBox, % InvoiceBalance
	
	return
}
*/

;CC 버튼 찾고 클릭
FindCCButtonAndClickIt(){

	Text:="|<CConLAMBS>*140$58.D002E0S005000882000M000UUE001UCQSLV0SQy0m+98409YM2DcYUE3YFU8UWG10GF50X2982194Hm7bYs7bYDU"

	if ok:=FindText(3016,149,150000,150000,0,0,Text)
	{
	  CoordMode, Mouse
	  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
	  MouseMove, X+W//2, Y+H//2
	  Click
	}
}

	
;지정된 창이 존재하는지 무한정 확인 후 그 창 활성화
CheckTheWindowPresentAndActiveIt(windowtitle){
	WinWait, % windowtitle
	WinActivate, % windowtitle
	return
}
	
	
	;위치 받아서 클릭하기
ClickAtThePoint(XPoint, YPoint){
	MouseClick, l, XPoint, YPoint, 1
	Sleep 1000
	return
}


	;상태바에서 알트키, 방향키 등 눌러서 Create Sales Orders Small Tab열기
	OpenCreateSalesOrdersSmallTab(){

		Start()		

		; 혹시 위의 동작으로도 Create Sales Orders Small 탭으로 넘어가지 않았을 때 백업 용도로 비활성화된 Create Sales Orders Small 버튼 클릭하기
		; 이 코드는 메뉴를 클릭해서 Create Sales Orders Small 을 새로 여는 게 아니라 이미 열려있는데 활성화 되지 않았을 때만 유효함
		; 혹시 마우스 커서 등이 가려서 못 찾는 것을 방지하기 위해 일단 마우스를 옮기기
;		MouseMove, A_ScreenWidth/2, A_ScreenHeight/2


		Text:="|<>*176$26.Q002M000a0009UhnmKAY4YG8792W22EcUYbm8D9U"
		
		if ok:=FindText(296,90,150000,150000,0,0,Text)
		{
			CoordMode, Mouse
			X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
			MouseMove, X+W//2, Y+H//2
			Click
				
			; 탭 버튼 클릭 후 화면의 정중앙으로 마우스 포인터 이동시키기
;			MouseMove, A_ScreenWidth/2, A_ScreenHeight/2
			;~ MsgBox
				
			; 버튼 찾아서 클릭했으면 그냥 함수를 빠져나가기
			;~ MsgBox, BEING FOUND BY FindText()
			return
		}
*/		

		Send, {Alt}
		Sleep 10
	
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10
		
		Send, {Right}
		Sleep 10

		Send, {Down}
		Sleep 10
		
		Send, {Right}
		Sleep 10

		Send, {Down}
		Sleep 10
		
		Send, {Enter}
		Sleep 100		

		return
	}








;===== Copy The Following Functions To Your Own Code Just once =====


; Note: parameters of the X,Y is the center of the coordinates,
; and the W,H is the offset distance to the center,
; So the search range is (X-W, Y-H)-->(X+W, Y+H).
; err1 is the character "0" fault-tolerant in percentage.
; err0 is the character "_" fault-tolerant in percentage.
; Text can be a lot of text to find, separated by "|".
; ruturn is a array, contains the X,Y,W,H,Comment results of Each Find.

FindText(x,y,w,h,err1,err0,text)
{
  xywh2xywh(x-w,y-h,2*w+1,2*h+1,x,y,w,h)
  if (w<1 or h<1)
    return, 0
  bch:=A_BatchLines
  SetBatchLines, -1
  ;--------------------------------------
  GetBitsFromScreen(x,y,w,h,Scan0,Stride,bits)
  ;--------------------------------------
  sx:=0, sy:=0, sw:=w, sh:=h, arr:=[]
  Loop, 2 {
  Loop, Parse, text, |
  {
    v:=A_LoopField
    IfNotInString, v, $, Continue
    Comment:="", e1:=err1, e0:=err0
    ; You Can Add Comment Text within The <>
    if RegExMatch(v,"<([^>]*)>",r)
      v:=StrReplace(v,r), Comment:=Trim(r1)
    ; You can Add two fault-tolerant in the [], separated by commas
    if RegExMatch(v,"\[([^\]]*)]",r)
    {
      v:=StrReplace(v,r), r1.=","
      StringSplit, r, r1, `,
      e1:=r1, e0:=r2
    }
    StringSplit, r, v, $
    color:=r1, v:=r2
    StringSplit, r, v, .
    w1:=r1, v:=base64tobit(r2), h1:=StrLen(v)//w1
    if (r0<2 or h1<1 or w1>sw or h1>sh or StrLen(v)!=w1*h1)
      Continue
    ;--------------------------------------------
    if InStr(color,"-")
    {
      r:=e1, e1:=e0, e0:=r, v:=StrReplace(v,"1","_")
      v:=StrReplace(StrReplace(v,"0","1"),"_","0")
    }
    mode:=InStr(color,"*") ? 1:0
    color:=RegExReplace(color,"[*\-]") . "@"
    StringSplit, r, color, @
    color:=Round(r1), n:=Round(r2,2)+(!r2)
    n:=Floor(255*3*(1-n)), k:=StrLen(v)*4
    VarSetCapacity(ss, sw*sh, Asc("0"))
    VarSetCapacity(s1, k, 0), VarSetCapacity(s0, k, 0)
    VarSetCapacity(rx, 8, 0), VarSetCapacity(ry, 8, 0)
    len1:=len0:=0, j:=sw-w1+1, i:=-j
    ListLines, Off
    Loop, Parse, v
    {
      i:=Mod(A_Index,w1)=1 ? i+j : i+1
      if A_LoopField
        NumPut(i, s1, 4*len1++, "int")
      else
        NumPut(i, s0, 4*len0++, "int")
    }
    ListLines, On
    e1:=Round(len1*e1), e0:=Round(len0*e0)
    ;--------------------------------------------
    if PicFind(mode,color,n,Scan0,Stride,sx,sy,sw,sh
      ,ss,s1,s0,len1,len0,e1,e0,w1,h1,rx,ry)
    {
      rx+=x, ry+=y
      arr.Push(rx,ry,w1,h1,Comment)
    }
  }
  if (arr.MaxIndex())
    Break
  if (A_Index=1 and err1=0 and err0=0)
    err1:=0.05, err0:=0.05
  else Break
  }
  SetBatchLines, %bch%
  return, arr.MaxIndex() ? arr:0
}

PicFind(mode, color, n, Scan0, Stride
  , sx, sy, sw, sh, ByRef ss, ByRef s1, ByRef s0
  , len1, len0, err1, err0, w, h, ByRef rx, ByRef ry)
{
  static MyFunc
  if !MyFunc
  {
    x32:="5589E583EC408B45200FAF45188B551CC1E20201D08945F"
    . "48B5524B80000000029D0C1E00289C28B451801D08945D8C74"
    . "5F000000000837D08000F85F00000008B450CC1E81025FF000"
    . "0008945D48B450CC1E80825FF0000008945D08B450C25FF000"
    . "0008945CCC745F800000000E9AC000000C745FC00000000E98"
    . "A0000008B45F483C00289C28B451401D00FB6000FB6C02B45D"
    . "48945EC8B45F483C00189C28B451401D00FB6000FB6C02B45D"
    . "08945E88B55F48B451401D00FB6000FB6C02B45CC8945E4837"
    . "DEC007903F75DEC837DE8007903F75DE8837DE4007903F75DE"
    . "48B55EC8B45E801C28B45E401D03B45107F0B8B55F08B452C0"
    . "1D0C600318345FC018345F4048345F0018B45FC3B45240F8C6"
    . "AFFFFFF8345F8018B45D80145F48B45F83B45280F8C48FFFFF"
    . "FE9A30000008B450C83C00169C0E803000089450CC745F8000"
    . "00000EB7FC745FC00000000EB648B45F483C00289C28B45140"
    . "1D00FB6000FB6C069D02B0100008B45F483C00189C18B45140"
    . "1C80FB6000FB6C069C04B0200008D0C028B55F48B451401D00"
    . "FB6000FB6C06BC07201C83B450C730B8B55F08B452C01D0C60"
    . "0318345FC018345F4048345F0018B45FC3B45247C948345F80"
    . "18B45D80145F48B45F83B45280F8C75FFFFFF8B45242B45488"
    . "3C0018945488B45282B454C83C00189454C8B453839453C0F4"
    . "D453C8945D8C745F800000000E9E3000000C745FC00000000E"
    . "9C70000008B45F80FAF452489C28B45FC01D08945F48B45408"
    . "945E08B45448945DCC745F000000000EB708B45F03B45387D2"
    . "E8B45F08D1485000000008B453001D08B108B45F401D089C28"
    . "B452C01D00FB6003C31740A836DE001837DE00078638B45F03"
    . "B453C7D2E8B45F08D1485000000008B453401D08B108B45F40"
    . "1D089C28B452C01D00FB6003C30740A836DDC01837DDC00783"
    . "08345F0018B45F03B45D87C888B551C8B45FC01C28B4550891"
    . "08B55208B45F801C28B45548910B801000000EB2990EB01908"
    . "345FC018B45FC3B45480F8C2DFFFFFF8345F8018B45F83B454"
    . "C0F8C11FFFFFFB800000000C9C25000"
    x64:="554889E54883EC40894D10895518448945204C894D288B4"
    . "5400FAF45308B5538C1E20201D08945F48B5548B8000000002"
    . "9D0C1E00289C28B453001D08945D8C745F000000000837D100"
    . "00F85000100008B4518C1E81025FF0000008945D48B4518C1E"
    . "80825FF0000008945D08B451825FF0000008945CCC745F8000"
    . "00000E9BC000000C745FC00000000E99A0000008B45F483C00"
    . "24863D0488B45284801D00FB6000FB6C02B45D48945EC8B45F"
    . "483C0014863D0488B45284801D00FB6000FB6C02B45D08945E"
    . "88B45F44863D0488B45284801D00FB6000FB6C02B45CC8945E"
    . "4837DEC007903F75DEC837DE8007903F75DE8837DE4007903F"
    . "75DE48B55EC8B45E801C28B45E401D03B45207F108B45F0486"
    . "3D0488B45584801D0C600318345FC018345F4048345F0018B4"
    . "5FC3B45480F8C5AFFFFFF8345F8018B45D80145F48B45F83B4"
    . "5500F8C38FFFFFFE9B60000008B451883C00169C0E80300008"
    . "94518C745F800000000E98F000000C745FC00000000EB748B4"
    . "5F483C0024863D0488B45284801D00FB6000FB6C069D02B010"
    . "0008B45F483C0014863C8488B45284801C80FB6000FB6C069C"
    . "04B0200008D0C028B45F44863D0488B45284801D00FB6000FB"
    . "6C06BC07201C83B451873108B45F04863D0488B45584801D0C"
    . "600318345FC018345F4048345F0018B45FC3B45487C848345F"
    . "8018B45D80145F48B45F83B45500F8C65FFFFFF8B45482B859"
    . "000000083C0018985900000008B45502B859800000083C0018"
    . "985980000008B45703945780F4D45788945D8C745F80000000"
    . "0E90B010000C745FC00000000E9EC0000008B45F80FAF45488"
    . "9C28B45FC01D08945F48B85800000008945E08B85880000008"
    . "945DCC745F000000000E9800000008B45F03B45707D368B45F"
    . "04898488D148500000000488B45604801D08B108B45F401D04"
    . "863D0488B45584801D00FB6003C31740A836DE001837DE0007"
    . "8778B45F03B45787D368B45F04898488D148500000000488B4"
    . "5684801D08B108B45F401D04863D0488B45584801D00FB6003"
    . "C30740A836DDC01837DDC00783C8345F0018B45F03B45D80F8"
    . "C74FFFFFF8B55388B45FC01C2488B85A000000089108B55408"
    . "B45F801C2488B85A80000008910B801000000EB2F90EB01908"
    . "345FC018B45FC3B85900000000F8C05FFFFFF8345F8018B45F"
    . "83B85980000000F8CE6FEFFFFB8000000004883C4405DC390"
    MCode(MyFunc, A_PtrSize=8 ? x64:x32)
  }
  return, DllCall(&MyFunc, "int",mode
    , "uint",color, "int",n, "ptr",Scan0, "int",Stride
    , "int",sx, "int",sy, "int",sw, "int",sh
    , "ptr",&ss, "ptr",&s1, "ptr",&s0
    , "int",len1, "int",len0, "int",err1, "int",err0
    , "int",w, "int",h, "int*",rx, "int*",ry)
}

xywh2xywh(x1,y1,w1,h1,ByRef x,ByRef y,ByRef w,ByRef h)
{
  SysGet, zx, 76
  SysGet, zy, 77
  SysGet, zw, 78
  SysGet, zh, 79
  left:=x1, right:=x1+w1-1, up:=y1, down:=y1+h1-1
  left:=left<zx ? zx:left, right:=right>zx+zw-1 ? zx+zw-1:right
  up:=up<zy ? zy:up, down:=down>zy+zh-1 ? zy+zh-1:down
  x:=left, y:=up, w:=right-left+1, h:=down-up+1
}

GetBitsFromScreen(x,y,w,h,ByRef Scan0,ByRef Stride,ByRef bits)
{
  VarSetCapacity(bits,w*h*4,0), bpp:=32
  Scan0:=&bits, Stride:=((w*bpp+31)//32)*4
  Ptr:=A_PtrSize ? "UPtr" : "UInt", PtrP:=Ptr . "*"
  win:=DllCall("GetDesktopWindow", Ptr)
  hDC:=DllCall("GetWindowDC", Ptr,win, Ptr)
  mDC:=DllCall("CreateCompatibleDC", Ptr,hDC, Ptr)
  ;-------------------------
  VarSetCapacity(bi, 40, 0), NumPut(40, bi, 0, "int")
  NumPut(w, bi, 4, "int"), NumPut(-h, bi, 8, "int")
  NumPut(1, bi, 12, "short"), NumPut(bpp, bi, 14, "short")
  ;-------------------------
  if hBM:=DllCall("CreateDIBSection", Ptr,mDC, Ptr,&bi
    , "int",0, PtrP,ppvBits, Ptr,0, "int",0, Ptr)
  {
    oBM:=DllCall("SelectObject", Ptr,mDC, Ptr,hBM, Ptr)
    DllCall("BitBlt", Ptr,mDC, "int",0, "int",0, "int",w, "int",h
      , Ptr,hDC, "int",x, "int",y, "uint",0x00CC0020|0x40000000)
    DllCall("RtlMoveMemory","ptr",Scan0,"ptr",ppvBits,"ptr",Stride*h)
    DllCall("SelectObject", Ptr,mDC, Ptr,oBM)
    DllCall("DeleteObject", Ptr,hBM)
  }
  DllCall("DeleteDC", Ptr,mDC)
  DllCall("ReleaseDC", Ptr,win, Ptr,hDC)
}

base64tobit(s)
{
  ListLines, Off
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  StringCaseSense, On
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:=(i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,A_LoopField,v)
  }
  StringCaseSense, Off
  s:=SubStr(s,1,InStr(s,"1",0,0)-1)
  s:=RegExReplace(s,"[^01]+")
  ListLines, On
  return, s
}

bit2base64(s)
{
  ListLines, Off
  s:=RegExReplace(s,"[^01]+")
  s.=SubStr("100000",1,6-Mod(StrLen(s),6))
  s:=RegExReplace(s,".{6}","|$0")
  Chars:="0123456789+/ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    . "abcdefghijklmnopqrstuvwxyz"
  SetFormat, IntegerFast, d
  Loop, Parse, Chars
  {
    i:=A_Index-1, v:="|" . (i>>5&1) . (i>>4&1)
      . (i>>3&1) . (i>>2&1) . (i>>1&1) . (i&1)
    s:=StrReplace(s,v,A_LoopField)
  }
  ListLines, On
  return, s
}

ASCII(s)
{
  if RegExMatch(s,"(\d+)\.([\w+/]{3,})",r)
  {
    s:=RegExReplace(base64tobit(r2),".{" r1 "}","$0`n")
    s:=StrReplace(StrReplace(s,"0","_"),"1","0")
  }
  else s=
  return, s
}

MCode(ByRef code, hex)
{
  ListLines, Off
  bch:=A_BatchLines
  SetBatchLines, -1
  VarSetCapacity(code, StrLen(hex)//2)
  Loop, % StrLen(hex)//2
    NumPut("0x" . SubStr(hex,2*A_Index-1,2), code, A_Index-1, "char")
  Ptr:=A_PtrSize ? "UPtr" : "UInt"
  DllCall("VirtualProtect", Ptr,&code, Ptr
    ,VarSetCapacity(code), "uint",0x40, Ptr . "*",0)
  SetBatchLines, %bch%
  ListLines, On
}

; You can put the text library at the beginning of the script,
; and Use Pic(Text,1) to add the text library to Pic()'s Lib,
; Use Pic("comment1|comment2|...") to get text images from Lib
Pic(comments, add_to_Lib=0) {
  static Lib:=[]
  if (add_to_Lib)
  {
    re:="<([^>]*)>[^$]+\$\d+\.[\w+/]{3,}"
    Loop, Parse, comments, |
      if RegExMatch(A_LoopField,re,r)
        Lib[Trim(r1)]:=r
  }
  else
  {
    text:=""
    Loop, Parse, comments, |
      text.="|" . Lib[Trim(A_LoopField)]
    return, text
  }
}


/***** C source code of machine code *****

int __attribute__((__stdcall__)) PicFind(int mode
  , unsigned int c, int n, unsigned char * Bmp
  , int Stride, int sx, int sy, int sw, int sh
  , char * ss, int * s1, int * s0
  , int len1, int len0, int err1, int err0
  , int w, int h, int * rx, int * ry)
{
  int x, y, o=sy*Stride+sx*4, j=Stride-4*sw, i=0;
  int r, g, b, rr, gg, bb, e1, e0;
  if (mode==0)  // Color Mode
  {
    rr=(c>>16)&0xFF; gg=(c>>8)&0xFF; bb=c&0xFF;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
      {
        r=Bmp[2+o]-rr; g=Bmp[1+o]-gg; b=Bmp[o]-bb;
        if (r<0) r=-r; if (g<0) g=-g; if (b<0) b=-b;
        if (r+g+b<=n) ss[i]='1';
      }
  }
  else  // Gray Threshold Mode
  {
    c=(c+1)*1000;
    for (y=0; y<sh; y++, o+=j)
      for (x=0; x<sw; x++, o+=4, i++)
        if (Bmp[2+o]*299+Bmp[1+o]*587+Bmp[o]*114<c)
          ss[i]='1';
  }
  w=sw-w+1; h=sh-h+1;
  j=len1>len0 ? len1 : len0;
  for (y=0; y<h; y++)
  {
    for (x=0; x<w; x++)
    {
      o=y*sw+x; e1=err1; e0=err0;
      for (i=0; i<j; i++)
      {
        if (i<len1 && ss[o+s1[i]]!='1' && (--e1)<0)
          goto NoMatch;
        if (i<len0 && ss[o+s0[i]]!='0' && (--e0)<0)
          goto NoMatch;
      }
      rx[0]=sx+x; ry[0]=sy+y;
      return 1;
      NoMatch:
      continue;
    }
  }
  return 0;
}

*/


;================= The End =================

;























Exitapp

Esc::
 Exitapp