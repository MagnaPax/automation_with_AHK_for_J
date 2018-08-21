#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#Include %A_ScriptDir%\lib\

#Include ChromeGet.ahk
#Include CommN41.ahk



aaa = Failed ASDF

if aaa contains Failed
{
	MsgBox, pending
}

MsgBox


LASMemo = Comment: ASDF

LASMemo := RegExReplace(LASMemo, "Comment: None(.*)", "$1")  ; $1 역참조를 사용하여 Comment: None 이외의 메모 내용이 있으면 변수에 저장

StringUpper, LASMemo, LASMemo ; 대문자로 바꾸기

MsgBox, % LASMemo















MsgBox





/*
; 사용자의 마우스 이동 막음
BlockInput, MouseMove

; 사용자의 마우스 이동 허용
BlockInput, MouseMoveOff
*/



/*
; 커서 상태가 작업처리중이면 끝날때까지 기다리기
while (A_cursor = "Wait")
	Sleep 3000
*/




/*
;  원하는 값을 못 찾았으면 ErrorLevel 통해서 if문 실행시키기
HTML_Source = 1234
HTML_Source = asdf

MsgBox, % HTML_Source

FoundPos := RegExMatch(HTML_Source, "1234")

MsgBox, % FoundPos

if ErrorLevel
{
	MsgBox, Not Found
}
*/



/*
StringUpper, Name, Name ; 대문자로 바꾸기
*/


/*
; https://autohotkey.com/boards/viewtopic.php?t=23286
; 배열에서 값을 찾아서 그 위치를 반환

HasVal(haystack, needle) {
	if !(IsObject(haystack)) || (haystack.Length() = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}
*/


/*
; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)
*/


/*
; 루프 언제 탈출하는지 확인할 수 있는 예제
Loop{ ; 1

	MsgBox, loop 1 out
	break

	Loop{ ; 2
		
		MsgBox, loop 2 out
		break
		
		Loop{ ; 3
			
			MsgBox, Loop 3 out
			break			
		}

	}

}

MsgBox, all loop out


*/


/*
; gui로 만든 프로그래스 바
TotalLoops = 57
Gui, -Caption +AlwaysOnTop +LastFound
Gui, Add, Text, x12 y9 w100 h20 , S E A R C H I N G . . .
Gui, Add, Progress, w410 Range0-%TotalLoops% cRed vProgress

;~ Gui, Show
Gui, Show, w437 h84, SEARCHING ITEMS


Loop, %TotalLoops%{
	
	GuiControl,,Progress, +1
	Sleep, 100
}

Gui Destroy
;~ GuiClose:
;~ Gui Destroy
;~ ExitApp

MsgBox pause

*/


/*
Array := object() ; 배열 선언
Array.Insert("B1129") ; 배열에 값 넣기
Array := [] ; 배열 초기화
*/


/*
; 화면에서 찾는 값이 여러개일 때 첫번째, 두번째 찾기
driver := ChromeGet()

if(driver.FindElementByXPath("//*[contains(text(), '" keyword "')]"))
{
	MsgBox, % driver.FindElementByXPath("(//*[contains(text(), '" keyword "')])").Attribute("innerText")
	MsgBox, % driver.FindElementByXPath("(//*[contains(text(), '" keyword "')])[2]").Attribute("innerText")
}

; 이건 배열에서 해당하는 것
Loop, % Array_AvailableDate_Sorted.MaxIndex(){
	MsgBox % "Element number " . A_Index . " is " . Array_AvailableDate_Sorted[A_Index]
	MsgBox, % driver.FindElementByXPath("//*[text() = '" Array_AvailableDate_Sorted[A_Index] "']//parent::td//child::a").Attribute("innerText") ; 해당하는 첫 번째 값
	MsgBox, % driver.FindElementByXPath("(//*[text() = '" Array_AvailableDate_Sorted[A_Index] "'])[2]//parent::td//child::a").Attribute("innerText") ; 해당하는 두 번째 값
}


; 아래처럼 해도 된다
Array_AvailableDate_Sorted := object()
Array_AvailableDate_Sorted.Insert("05/10/2018")
driver := ChromeGet()
i = 2
Value := Array_AvailableDate_Sorted[A_Index]
Xpath = (//*[text() = '%Value%'])[%i%]//parent::td//child::a

Loop, % Array_AvailableDate_Sorted.MaxIndex(){
	MsgBox, % driver.FindElementByXPath(Xpath).Attribute("innerText")
}



*/



/*
; 배열에서 값을 찾아서 그 위치를 반환
arr := ["a", "b", "", "d"]

MsgBox % HasVal(arr, "a") "`n"    ; return 1
       . HasVal(arr, "e") "`n"    ; return 0
       . HasVal(arr, "d")         ; return 4

HasVal(haystack, needle) {
	if !(IsObject(haystack)) || (haystack.Length() = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}
*/

/*
https://www.google.com/search?newwindow=1&source=hp&ei=kIGDWpXYLdLajwOtnLT4DA&q=autohotkey+array+sort&oq=autohotkey+array+sort&gs_l=psy-ab.3..0i22i30k1.3181.7947.0.8239.32.20.0.0.0.0.148.1671.8j8.17.0....0...1.1.64.psy-ab..19.13.1372.6..0j35i39k1j0i20i264k1j0i203k1.104.qKKh-koFig0	
; 배열 작은 값부터 차례대로 정렬
QS := new Quickselect
Loop, % Arr.MaxIndex()
	MsgBox, % QS.Select(Arr, A_Index)
*/


/*
; 화면에서 PONumber 찾기
if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
	driver.FindElementByXPath("//*[text() = '" PONumber "']").click()

Xpath = //*[text() = '%PONumber%']
if(driver.FindElementByXPath(Xpath))
	driver.FindElementByXPath(Xpath).click()
*/

/*
; REFRESH THIS PAGE
driver.refresh()
*/

/*
; 브라우저 닫기
;~ driver.quit() ; 이거 작동하지 않음 closing all the browsers
driver.close() ; closing just one tab of the browser
*/


/*
; PO Number 링크 나타날때까지 기다림
Xpath = //*[text() = '%PONumber%']
Sleep 500
while(!driver.FindElementByXPath(Xpath).Attribute("innerText"))
	Sleep 100

Sleep 1000
; PO Number 검색한 현재 화면에서 정확히 원하는 PO Number만 찾아서 클릭하기
; 예를들어 MTR1D39747D26 로 검색했으면 검색된 화면에는 MTR1D39747D26 뿐만 아니라 MTR1D39747D26-BO1 등도 같이 표시되기 때문에 딱 원하는 PO Number만 클릭하기 위해
if(driver.FindElementByXPath("//*[text() = '" PONumber "']"))
	driver.FindElementByXPath("//*[text() = '" PONumber "']").click()
*/

/*
; Element 가 있으면 if문 실행
if(driver.FindElementByXPath(Xpath).isDisplayed()){
	MsgBox, DISPLAYED
}
*/

/*
; Element 가 없으면 if문 실행
if(!driver.FindElementByXPath(Xpath).isDisplayed()){
	MsgBox, NOT DISPLAYED
}
*/

/*
; Element 가 화면에 표시될때까지 기다린 후 클릭하기
loop{
	if(driver.FindElementByXPath(Xpath).isDisplayed()){
		driver.FindElementByXPath(Xpath).click()
		break
	}
	Sleep 100
}
*/

/*
; Element 가 화면에 표시됐는지 알아보는 코드
driver.FindElementByXPath(Xpath).isDisplayed()
*/

/*
; element status 어떤지 알아보는 함수. 0을 반환하면 문제 있는 것
MsgBox, % driver.FindElementByXPath(Xpath).isEnabled()
*/


/*
; 라디오 버튼 중 선택된 버튼이 어떤 것인지 나타내기(if 쓰면 그에 맞는 동작도 시킬 수 있겠지)
action:=["Nothing", "Something", "Everything possible", "In Jesus"]
    
Gui, New
Gui, Add, Text,, What should I do?
Gui, Add, Radio, vRadioGroup, % action[1]
Gui, Add, Radio,, % action[2]
Gui, Add, Radio,, % action[3]
Gui, Add, Radio,, % action[4]
Gui, Show
Return
    
Guiclose:
Gui, Submit
MsgBox % "I'll do " action[RadioGroup]
exitApp
return
*/

/*
; 파일을 삭제
FileDelete, C:\temp files\*.tmp
*/

/*
; 파일이나 폴더 있는지 확인
IfExist, D:\
    MsgBox, 드라이브가 존재합니다.
IfExist, D:\Docs\*.txt
    MsgBox, 적어도 하나의 .txt 파일이 존재합니다.
IfNotExist, C:\Temp\FlagFile.txt
    MsgBox, 목표 파일이 존재하지 않습니다.
*/


/*
; CustomerNoteOnWebVal 변수 안에 있는 메모 내용 CustomerNoteOnWebVal.txt 파일에 저장하기
FileAppend, %CustomerNoteOnWebVal%, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt
*/	

/*
; CustomerNoteOnWebVal.txt 파일을 EmptyFile.txt 로 덮어씌워 초기화 하기
FileCopy, %A_ScriptDir%\CreatedFiles\EmptyFile.txt, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt, 1
*/

/*
; CustomerNoteOnWebVal.txt 내용을 CustomerNoteOnWebVal 변수에 저장하기
FileRead, CustomerMemoOnLAMBS, %A_ScriptDir%\CreatedFiles\CustomerNoteOnWeb.txt
*/

/*
OutArray := StrSplit(Input, ",")  ; 콤마 나올때마다 문자열 나누기
Loop % OutArray.Maxindex(){
	OutArray[A_Index] := Trim(OutArray[A_Index])
	MsgBox % "Element number " . A_Index . " is |||" . OutArray[A_Index] . "|||"
}
*/

/* 알파벳과 숫자만 저장 (알페벳과 숫자 제외한 모든 것을 "" 로 바꿈. 즉, 삭제)
HumanVerificationCode := RegExReplace(HumanVerificationCode, "[^a-zA-Z0-9]", "")
*/

/*
Data = abc123123
NewStr := RegExReplace(Data, "123$", "xyz")  ; 맨 마지막 123만 xyz로 바꿔서 변수에 "abc123xyz" 저장. abcxyzxyz가 안 되고 맨 끝에만 바뀌는 이유는 $는 끝에만 부합을 허용하기 때문에
*/

/*
NewStr := RegExReplace("abcXYZ123", "abc(.*)123", "aaa$1zzz")  ; $1 역참조를 사용하여 "aaaXYZzzz"을 돌려 줍니다.
NewStr := RegExReplace("abcXYZ123", "abc(.*)123", "$1")  ; $1 역참조를 사용하여 "XYZ"을 돌려 줍니다.
NewStr := RegExReplace("abcXYZ123", "bc(.*)123", "$1")  ; $1 역참조를 사용하여 "aXYZ"을 돌려 줍니다.
MsgBox, % NewStr
*/

/*
; IfEqual 결과 어떻게 나오는지
URLofCustPO = 1234
CurrentURL = 1234

IfNotEqual, URLofCustPO, %CurrentURL%
{
	MsgBox, IT'S NOT EQUAL
}
IfEqual, URLofCustPO, %CurrentURL%
{
	MsgBox, IT'S EQUAL
}

;~ if CurrentURL contains URLofCustPO
if(CurrentURL != URLofCustPO)
{
	MsgBox, % "not matched"
}

MsgBox
*/


/*
; 소스에서 숫자만 추출해서 특정 스트링 뒤에 붙여서 url 만들기
SourceStr = <a _ngcontent-c10="" href="#/order/13371385">MTR1F21F7899C</a>
UnquotedOutputVar = .*/order/(\d*).*>.*

RefinedStr := RegExReplace(SourceStr, UnquotedOutputVar, "$1")  ; UnquotedOutputVar 조건을 보고 $1 역참조를 사용하여 숫자만 RefinedStr 변수에 저장
AddedFinalURL = https://vendoradmin.fashiongo.net/#/order/

URLofCustPO := AddedFinalURL . RefinedStr

MsgBox, % "[SourceStr]`n" . SourceStr . "`n`n`n[URLofCustPO]`n" . URLofCustPO
*/


/*
; RegExMatch 사용해서 문자열 추출하기
; RegExMatch 명령어 사용 시 첫 번째와 두 번째의 괄호 차이에 따른 SubPat2 결과값의 변화, 세 번째와 네 번째의 imU) 유무에 따른 결과값의 변화를 보라
FileRead, Source, %A_ScriptDir%\RegExMatchSample[Do Not Delete].txt ; 공백과 여러 쓸데없는 문자열이 포함된 소스를 읽어서 Source 변수에 저장
MsgBox, % Source


; 이 동작에서는 조건식을 전체 괄호 한 뒤 필요한 부분에 또 괄호를 했을 때 원하는 결과값이 SubPat2 변수에 저장됐다
RegExMatch(Source, "imU)Via\](.*)\[Ship", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2

RegExMatch(Source, "imU)(Via\](.*)\[Ship)", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2


; 이 동작에서는 앞에 imU)를 붙였을 때는 조건식 다음부터 문자열 끝까지의 결과가 추출이 안 되다가 없애니 추출됐다
RegExMatch(Source, "imU)(Via\].*Memo](.*))", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2

RegExMatch(Source, "(Via\].*Memo](.*))", SubPat)
MsgBox, % "SubPat`n" . SubPat . "`n`n`nSubPat2`n" . SubPat2
*/


/* 항상 위 메세지 Ok
MsgBox, 262144, Title, Message Placed Here
*/

/* 항상 위 메세지 Yes or No
MsgBox, 4100, Wintitle, Click Ok to continue
*/

/*
	;StringSplit 예제. word_array의 쓰임
	TestString = This is a test.
	; 공란(스페이스)이나 콤마가 나올때마다 나누고 마침표(.)는 제외해서 word_array에 저장
	StringSplit, word_array, TestString, `,|%A_Space%, .  ; 점은 제외합니다.

	MsgBox, The 4th word is %word_array4%.

	Colors = red,green,blue
	StringSplit, ColorArray, Colors, `,
	Loop, %ColorArray0%
	{
		this_color := ColorArray%a_index%
		MsgBox, Color number %a_index% is %this_color%.
	}
*/	



/*	
	; 메소드 예제
	; 배열에 들어있는 값 갯수 구하는 메소드
	; arr[0]에 111이 들어가고 arr[10]에 222가 들어가서 배열에 들어간 값 갯수는 2개
	array := Object() 
	array.length := "array_length" 
	array_length(object) 
	{ 
		  current_length := "0" 

		  loop_count := object.maxIndex() + 1 
		  loop % loop_count 
		  { 
				  current_index := a_index - 1 
				  if(object[current_index] != "") 
				  { 
						current_length++ 
				  } 
		  } 

		  return current_length 
	} 


	arr := Object("base", array)              ;==========> 객체 생성시에 다음과 같이 하면 상부에 정의된 length 를 사용할 수 있습니다. 
	arr[0] := 111 
	arr[10] := 222 
	Msgbox % arr.length()
*/


/*
	; 마우스 커서의 현재 위치를 실시간으로 열람
;	MouseGetPos, xpos, ypos
;	Msgbox, The cursor is at X%xpos% Y%ypos%. 

	; 이 예제에서 마우스를 이동시켜서 현재 마우스 아래에 있는
	; 창의 제목을 볼 수 있습니다:
	#Persistent
	SetTimer, WatchCursor, 100
	return

	WatchCursor:
	MouseGetPos, , , id, control
	MouseGetPos, , , , hWnd, 2
	WinGetTitle, title, ahk_id %id%
	WinGetClass, class, ahk_id %id%
	ToolTip, ahk_id(WinID):  %id%`nahk_class:  %class%`nWindow_Title:  %title%`nControl(ClassNN):  %control%`nhWnd:  %hWnd%
	return
*/

/*
q:: ;get control information for the active window (ClassNN and text)
WinGet, vCtlList, ControlList, A
vOutput := ""
Loop, Parse, vCtlList, `n
{
	vCtlClassNN := A_LoopField
	ControlGetText, vText, % vCtlClassNN, A
	vOutput .= vCtlClassNN "`t" vText "`r`n"
}

Clipboard := vOutput
MsgBox, % "done"
return
*/


/*
q::
title = ahk_class Notepad
Loop{
ControlGetText, OutputVar, Edit1, ahk_class FNWND3126
ControlSend, Edit1, %A_Index%, %title%
;~ ControlGetText, A, , Edit66, ahk_class FNWND3126
;~ Send, %a_index%
Sleep 1000
}
*/


/*
; 예제 #4: 실시간으로 활성 창의 콘트롤 리스트를 보여줍니다:
#Persistent
SetTimer, WatchActiveWindow, 200
return
WatchActiveWindow:
WinGet, ControlList, ControlList, A
ToolTip, %ControlList%
return
*/

	
/*
	; 다음의 작동하는 예제는 계속 갱신하면서
	; 현재 마우스 아래의 콘트롤의 위치와 이름을 보여줍니다:
	Loop
	{
		Sleep, 100
		MouseGetPos, , , WhichWindow, WhichControl
		ControlGetPos, x, y, w, h, %WhichControl%, ahk_id %WhichWindow%
		ToolTip, %WhichControl%`nX%X%`tY%Y%`nW%W%`t%H%
	}
*/			
		

/*
	;  크롬 새창에서 열기

	url = http://vendoradmin3.fashiongo.net/OrderDetails.aspx?po=MTR1BE9C4B535

	run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) url

	return
*/



/*
웹페이지 자동 로그인 
Loginname = user name
Password = pass word
URL = www.google.com

WB := ComObjCreate("InternetExplorer.Application")
WB.Visible := True
WB.Navigate(URL)
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10

wb.document.getElementById("login").value := Loginname
wb.document.getElementById("password").value := Password
wb.document.getElementsByTagName("Button")[1].Click()

While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10
Msgbox, Something like that i hope!
return
*/


/*
; 크롬 현재 창 url 얻기
hwndChrome := WinExist("ahk_class Chrome_WidgetWin_1")
AccChrome := Acc_ObjectFromWindow(hwndChrome)
AccAddressBar := GetElementByName(AccChrome, "Address and search bar")
MsgBox % AccAddressBar.accValue(0)

GetElementByName(AccObj, name) {
   if (AccObj.accName(0) = name)
      return AccObj
   
   for k, v in Acc_Children(AccObj)
      if IsObject(obj := GetElementByName(v, name))
         return obj
}
*/


/*
; Trim 함수 적용 예
; 문자열 양쪽이나 왼쪽 오른쪽의 문자를 없앤다
text := "  <-Here is Space / Here is Tap ->	"
MsgBox % "No trim:`t '" text "'"
    . "`nTrim:`t '" Trim(text) "'"
    . "`nLTrim:`t '" LTrim(text) "'"
    . "`nRTrim:`t '" RTrim(text) "'"
MsgBox % LTrim("00000123","0")
*/

/*
; continue 예제
; 이 예제는 5개의 MsgBox를 보여줍니다. 각각 6부터 10까지 담고 있습니다.
; 회돌이의 앞쪽 5회에, "continue" 명령어 때문에
; 회돌이가 MsgBox 줄에 도착하기 전에 다시 시작하는 것을 주목하십시오.
Loop, 10
{
    if A_Index <= 5
        continue
    MsgBox %A_Index%
}
*/


/*
; continue 예제
; 이 예제는 5개의 MsgBox를 보여줍니다. 각각 6부터 10까지 담고 있습니다.
; 회돌이의 앞쪽 5회에, "continue" 명령어 때문에
; 회돌이가 MsgBox 줄에 도착하기 전에 다시 시작하는 것을 주목하십시오.
Loop, 10
{
	if A_Index >= 5
	{
		MsgBox, %A_Index% oo
        continue
	}
    MsgBox %A_Index%
}
*/


/*  커서 상태가 작업처리중이면 끝날때까지 기다리기
Loop{
	if(A_cursor = "Wait"){
		Sleep 1000
	}
	else
		break
}

while (A_cursor = "Wait")
	Sleep 2000	
*/


/* 배열 여러개 리턴하는 예제
;~ https://stackoverflow.com/questions/5760058/how-to-return-multiple-arrays-from-a-function-in-javascript

Test := returnMultipleArrays()
MsgBox, % Test[1][1] "`, " Test[1][2] . "`n" . Test[2][1] "`, " Test[2][2] . "`n" . Test[3][1] "`, " Test[3][2]


returnMultipleArrays()
{
 Array1 := ["1", "2"]
 Array2 := ["3", "4"]
 Array3 := ["5", "6"]
 return [Array1, Array2, Array3]
}

*/

/* 다차원 배열 리턴
;~ https://autohotkey.com/board/topic/127830-a-newbs-request-for-help-2d-array/

MyArray := Basic2D()
MsgBox, % MyArray[1, 1]
return

Basic2D() {
    Arr := [] 
    Arr[0, 0] := "0,0"
    Arr[0, 1] := "0,1"
    Arr[1, 0] := "1,0"
    Arr[1, 1] := "1,1"
    return Arr
}
*/


/* 배열로부터 읽기 첫 번째 방법
Array:=[1,3,"ㅋㅋ"]
Loop % Array.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
}
*/


/* 배열로부터 읽기 두 번째 방법
Array:=[1,3,"ㅋㅋ"]
for index, element in Array
{
	MsgBox % "Element number " . index . " is " . element
}
*/


/*
	; 배열에 들어있는 값 갯수 구하는 함수
	; arr[0]에 111이 들어가고 arr[10]에 222가 들어가서 배열에 들어간 값 갯수는 2개
	arr := Object() 
	arr[0] := "a" 
	arr[12] := "b" 
	MsgBox, % Obj_Length(arr) ; 2출력 
	return 
	
	Obj_Length(obj) { 
		length := 0 
		for idx in obj 
			length++ 
		return length 
	}
*/




/*
Test := returnMultipleArrays()
MsgBox, % Test[1][1] "`, " Test[1][2] . "`n" . Test[2][1] "`, " Test[2][2] . "`n" . Test[3][1] "`, " Test[3][2]


returnMultipleArrays()
{
	
 Array1%1% := ["1"]
 Array2%1% := ["3"]
 Array3%1% := ["5"]
 	
 Array1%2% := ["2"]
 Array2%2% := ["4"]
 Array3%2% := ["6"]
 
 return [Array1, Array2, Array3]
}
*/





		
		
		
Esc::
 Exitapp
 Reload		