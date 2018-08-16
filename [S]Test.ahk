#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)


#Include %A_ScriptDir%\lib\

#Include FG.ahk
;~ #Include FG-1.ahk
#Include LAS.ahk
#Include ChromeGet.ahk ; 이미 열려있는 창을 사용할 수 있게 해주는 함수. 
#Include OOP.ahk
#Include CommWeb.ahk

global Array_StyleNo, Array_StylyColor




/*
Array_StyleNo := object()
Array_StylyColor := object()

;~ Array_StyleNo.Insert("B1059")
Array_StyleNo.Insert("")
Array_StyleNo.Insert("B1129")
Array_StyleNo.Insert("B1272")
Array_StyleNo.Insert("H1207")


;~ Array_StylyColor.Insert("BLACK")
Array_StylyColor.Insert("")
Array_StylyColor.Insert("ALMOND MIX")
Array_StylyColor.Insert("OFF WHITE MIX")
Array_StylyColor.Insert("BLACK")
*/


Array_StyleNo := object()
Array_StylyColor := object()

Array_StyleNo.Insert("")
Array_StyleNo.Insert("B1129")
Array_StyleNo.Insert("H1207")
Array_StyleNo.Insert("ABCD")

Array_StylyColor.Insert("")
Array_StylyColor.Insert("ALMOND MIX")
Array_StylyColor.Insert("BLACK")
Array_StylyColor.Insert("LALALA")


MsgBox, % Array_StyleNo


Str_Temp := Object()
Str_CheckUnUpdatedItems := object()


		; 업데이트가 안된 아이템이 있는지 확인하기
		; 정상적으로 모두 업데이트 되었으면 Array_StyleNo 배열에는 값이 없어야 된다
		Loop % Array_StyleNo.Maxindex(){
			;~ MsgBox % A_Index " PONumber is`n" . Array_StyleNo[A_Index]
			
			UnUpdatedItem_# := Array_StyleNo[A_Index]
			UnUpdatedItem_Color := Array_StylyColor[A_Index]
			
			; 만약 UnUpdatedItem_# 변수에 무언가 값이 있으면 업데이트가 안 된 아이템이 있다는 뜻
			if(UnUpdatedItem_# != ""){
				
				; 업데이트 안된 스타일 번호와 색깔을 최종적으로 Str_CheckUnUpdatedItems 배열에 넣기 위해 임시로 Str_Temp 배열에 값 넣기
				Str_Temp[1] := UnUpdatedItem_#
				Str_Temp[2] := UnUpdatedItem_Color				
				
				
				Str_CheckUnUpdatedItems.Insert(Str_Temp) ; Str_CheckUnUpdatedItems 안에 차곡차곡 넣기. Str_CheckUnUpdatedItems 배열은 다중 배열이 됨
				Str_Temp := [] ; Str_CheckUnUpdatedItems 안에 넣을때마다 값이 중복되지 않게 Str_Temp 배열 초기화 해주기
				
				MsgBox, 262144, Alert, %UnUpdatedItem_#% %UnUpdatedItem_Color% `nIS NOT UPDATED, PLEASE CHECK IT.
			}
		}

		
		UnUpdatedItemExists := Str_CheckUnUpdatedItems[1][1] ; 업데이트 안 된 아이템이 있는지 확인하기 위해
		#ofUnUpdatedItems := Str_CheckUnUpdatedItems.Maxindex() ; 업데이트 안 된 아이템 숫자
		
		MsgBox, % UnUpdatedItemExists
		
		; 만약 UnUpdatedItemExists 안에 값이 있으면 업데이트 안 된 아이템이 있다는 뜻
		if(UnUpdatedItemExists != ""){
			
			i = 1
			Loop % Str_CheckUnUpdatedItems.Maxindex(){
				
				UnUpdatedStyleNo := Str_CheckUnUpdatedItems[i][1]
				UnUpdatedStyleColor := Str_CheckUnUpdatedItems[i][2]
				
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				;~ MsgBox % "Customer PO : " PONumber . "n`n`Invoice No : " . InvoiceNo . "`n`n`n`n" . Str_CheckUnUpdatedItems.Maxindex() . " ITEMS ARE NOT UPDATED" . "`n`n" . "Element number " . A_Index . " is " . "`n" . Str_CheckUnUpdatedItems[i][1] . "  " . Str_CheckUnUpdatedItems[i][2]
				MsgBox, 262144, Alert, Customer PO : %PONumber%`n`nInvoice No : %InvoiceNo%`n`n%#ofUnUpdatedItems% ITEMS ARE NOT UPDATED `n`n`n`nElement number %A_Index% is `n`n%UnUpdatedStyleNo%  %UnUpdatedStyleColor%

				i++			
			}
		}







MsgBox, % Array_StyleNo.MaxIndex()



/* 배열로부터 읽기 첫 번째 방법
Loop % Array.Maxindex(){
	MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
}
*/

/* 배열로부터 읽기 두 번째 방법
for index, element in Array
{
	MsgBox % "Element number " . index . " is " . element
}
*/

/* 배열로부터 읽기
Array:=[1,3,"ㅋㅋ"]
Loop,% Array.MaxIndex()
	Msgbox,% Array[a_index]
*/










driver := New FG

ShippingFee = 13
InvoiceNo = 00000
TrackingNo = ABCDEFG



/* 뉴오더 처리
PONumber = MTR2FC22A95D
;~ PONumber = MTR1D39747D26 ; 뉴오더에 없는 것
AlreadyProcessedPONumberOrNot := driver.NewOrderProcessing(PONumber)
MsgBox, AlreadyProcessedPONumberOrNot : %AlreadyProcessedPONumberOrNot%
driver.ClosingCurrentlyOpenedBrowser()

MsgBox, method out

*/




/* 새창 열기
URL = https://vendoradmin.fashiongo.net/#/order/orders/new
driver.OpenNewBrowser(URL)
MsgBox
*/



; 현재 열린 창 닫기
;~ driver.ClosingCurrentlyOpenedBrowser()



; FG에 로그인 하기
;~ driver.Login()




;~ /*
; 오더창 열어서 열린 페이지의 PO 번호들 얻기(아이템 처리 위함)
Array_PONumber := object()
PONumber = MTR1D39747D26
Array_PONumber := driver.OpenOrderPageForItemUpdate(PONumber, Array_PONumber)


		Loop % Array_PONumber.Maxindex(){
			;~ MsgBox % "Element number " . A_Index . " is " . Array_PONumber[A_Index]
			MsgBox % A_Index " PONumber is`n" . Array_PONumber[A_Index]
		}

MsgBox, pause



; 아이템 업데이트 하기
; 배열을 먼저 읽어서 값이 들어있을때만 아래 메소드 호출하기. 아무런 값도 없으면 아예 호출하지 않기(왜냐면 일단 호출하면 
Loop % Array_PONumber.Maxindex(){
	
	; OpenOrderPageForItemUpdate 메소드를 한 번 실행하고 나오면 driver 인스턴스(객체)값이 소멸되기 때문에 다시 선언해줘야 됨
	driver := New FG

	MsgBox % A_Index " PONumber is`n" . Array_PONumber[A_Index]
	
	PONumber := Array_PONumber[A_Index]
	driver.UpdateItems(Array_StyleNo, Array_StylyColor, ShippingFee, InvoiceNo, TrackingNo, PONumber)

}


MsgBox, method out



		Loop % Array_StyleNo.Maxindex(){
			Loop % Array_StylyColor.Maxindex(){

				MsgBox % "Style No. " . A_Index . " is " . Array_StyleNo[A_Index]  "`n"  "Style Color" . A_Index . " is " . Array_StylyColor[A_Index]
			}
		break
		}



MsgBox, out

*/





/* 같은 Xpath 를 공유하는 2번째 3번째 값 사용하기

URL = https://vendoradmin.fashiongo.net/#/order/12792823
driver.OpenNewBrowser(URL) ; 테스트를 위해 코멘드 라인이 첨가된 크롬창 열기

driver := ChromeGet()
MsgBox, % driver.Window.Title "`n" driver.Url
loop, 3{
	
	;~ Xpath = //*[text() = '%StyleNo%']
	
			;~ Xpath2 = //a[contains(text(),'B1504')]//parent::div//parent::td//following-sibling::td//child::div[@class='text-s']
			;~ StyleColor := driver.FindElementByXPath(Xpath2).Attribute("innerText")
			
			MsgBox, % driver.FindElementByXPath("(//a[contains(text(),'B1504')]//parent::div//parent::td//following-sibling::td//child::div[@class='text-s'])").Attribute("innerText")
			MsgBox, % driver.FindElementByXPath("(//a[contains(text(),'B1504')]//parent::div//parent::td//following-sibling::td//child::div[@class='text-s'])[2]").Attribute("innerText")
			
			i := % A_Index
			Xpath = (//a[contains(text(),'P2390')]//parent::div//parent::td//following-sibling::td//child::div[@class='text-s'])[%i%]
			MsgBox, % driver.FindElementByXPath(Xpath).Attribute("innerText")
}
MsgBox
*/


/* 배열로부터 읽기
SubPat1 = asdf
Array := object()
Array.Insert(SubPat1)
;~ Array:=[1,3,"ㅋㅋ"]
MsgBox, % "The method value of Array.MaxIndex is " . Array.MaxIndex()
Loop,% Array.MaxIndex()
	Msgbox,% Array[a_index]

MsgBox, % "Value in Array is : " . Array[1]
*/

/*
i = 1
loop{
	MsgBox, % a_index
	if(i == 3){
		MsgBox, break
		break
	}
	i++
}
MsgBox, out
*/


		; 현재 페이지 제어하기 위한 메소드 호출
		;~ driver := ChromeGet()
		;~ MsgBox, % driver.Window.Title "`n" driver.Url
		

;~ /* 페이지에 있는 정보(Style No, Style Color, Pre Order 여부, cc번호, Billing Add, Shipping Add) 얻기

;~ https://vendoradmin.fashiongo.net/#/order/12855747
;~ https://vendoradmin.fashiongo.net/#/order/12551798

Arr_StyleNo := object()
Arr_#ofColors := object()

PONumber = MTR1DD95515B8 ; 프리오더값 있음
PONumber = MTR1CAEB32DC7-BO1-BO2-BO2 ; 백오더값 있음

;~ driver.GettingInfoFromCurrentPage(PONumber, Arr_StyleNo, Arr_#ofColors)
driver.GettingInfoFromCurrentPage(PONumber)

*/





/* LA_OpenNewBrowser 함수 열어서 새로운 브라우저 연 뒤 다음 처리를 해야되는데 처음 브라우저 시작할 때 첫 페이지 로딩 시간이 너무 오래 걸려서 이후의 과정이 꼬이고 있음. 
;  새로운 창을 열때 객체 선언 뒤 코멘드 라인을 추가해서 연 뒤 기다리는 코드를 이용할까 했는데 그럼 객체가 없어지면 브라우저 창도 없어지니까
la_driver := new LA

;~ URL = https://admin.lashowroom.com/index.php ; 뉴오더 검색창 주소
;~ la_driver.LA_OpenNewBrowser(URL)
;~ la_driver.LA_Login()
;~ MsgBox


PONumber = OP112436997
la_driver.LA_UpdateItems(PONumber)
*/


/* OOP 테스트.
driver := New OOP
;~ driver.AddingValuesInArray() ; 클래스 안에서 선언된 전역변수 배열에 값 집어넣기 (this 사용)
;~ driver.ReadingValuesFromArray() ; 위 메소드에서 배열에 입력된 값들 출력하기 (this 사용)

;~ driver.InsertValueInTheVariable() ; this 사용 예제
;~ driver.ReadingValueInTheVariable() ; this 사용 예제

driver := new ExtensTest
driver.Overriding_OriginalValueIs123() ; Overriding 예제 (상속)
*/











Exitapp

Esc::
Exitapp