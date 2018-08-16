#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Include %A_ScriptDir%\lib\

#Include ChromeGet.ahk
#Include T_Up.ahk

; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)



driver := ChromeGet()
Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody[1]/tr/td[4]/table/tbody/tr[2]/td[7]/div/label/div/div
Xpath = /html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody[1]/tr/td[4]/table/tbody/tr[2]/td[7]/div/label/div

if(!driver.FindElementByXPath(Xpath).isSelected())
	MsgBox, check box


vv := driver.FindElementByXPath(Xpath).isSelected()

MsgBox, % vv

MsgBox


Array_StyleNo := object()
Array_StylyColor := object()
Array_StylyQty := object()


/*
; MTR1E393150C3-BO1 테스트용
; Cancelled by Buyer 로 바꿔야 됨
; 백오더 보내려고 체크되어 있지 않았는데 체크해제 하려고 체크하는지 테스트용
Array_StyleNo.Insert("P2104-3")
Array_StylyColor.Insert("PEACH")
Array_StylyQty.Insert("6")
*/



;~ /*
; MTR1EB058B4E2 테스트용
Array_StyleNo.Insert("B1122")
Array_StylyColor.Insert("CORAL")
Array_StylyQty.Insert("6")
Array_StyleNo.Insert("")
Array_StylyColor.Insert("")
Array_StylyQty.Insert("")
Array_StyleNo.Insert("J1797")
Array_StylyColor.Insert("OFF WHITE")
Array_StylyQty.Insert("6")
*/





/*
Array_StyleNo.Insert("")
Array_StyleNo.Insert("B1715")
Array_StyleNo.Insert("P1111")
Array_StyleNo.Insert("P2222")
Array_StyleNo.Insert("B1814")
Array_StyleNo.Insert("B1741")
Array_StyleNo.Insert("J1962")



Array_StylyColor.Insert("")
Array_StylyColor.Insert("DENIM BLUE MIX")
Array_StylyColor.Insert("LALALA")
Array_StylyColor.Insert("HAHAHA")
Array_StylyColor.Insert("DUSTY PINK")
Array_StylyColor.Insert("BLACK")
Array_StylyColor.Insert("ROYAL BLUE")



Array_StylyQty.Insert("")
Array_StylyQty.Insert("6")
Array_StylyQty.Insert("6")
Array_StylyQty.Insert("6")
Array_StylyQty.Insert("6")
Array_StylyQty.Insert("6")
Array_StylyQty.Insert("6")
*/





/*

MTR1E6AFD7C28

MTR1EDDFE4DD3 ; 이건 색깔이 3개라 체크박스도 3개가 있음

MTR1EECF24383 ; 인스탁 아이템 없고 프리오더 아이템만 있음

MTR1E7E3FCE48 ; Import 아이템 두 개

MTR1E73784A7E ; 수입 아이템 한 개

MTR1E393150C3-BO1 ; 수입과 국내가 각각 한 개씩. 모두 날짜 없음

MTR1EB058B4E2 ; 국내와 수입품이 섞여 있음. 모든 아이템에 날짜 없음


MTR1E640278CC ; cancelled by vendor 오더. 수입과 미국내 아이템이 섞여 있음


아이템들의 체크 박스들
/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody//child::div[@class='check-square']

아이템 번호와 프리오더 날짜가 같이 선택되는 xpath
/html/body/fg-root/div[1]/fg-secure-layout/div/div[2]/fg-order-detail/div[4]/div[2]/div[2]/table/tbody//child::div[@class='check-square']//ancestor::tbody//child::div[@class='order-table__no']

소스코드에서 체크박스 찾기
<div _ngcontent-c7="" class="check-square"><div
*/


; ##########
; FG PA 는 어떻게 할까
; ##########


MsgBox


F_driver := new FG
;~ F_driver := new FG_fake

; 처음엔 Domestic 오더 처리가 아니니 거짓인 0값으로 시작
IsItSecondCallToProcessDomesticItems = 0
;~ F_driver.ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems)
F_driver.ProcessingOfItemsAppearAsInStockButBeingKickedAsBoItems(IsItSecondCallToProcessDomesticItems, Array_StyleNo, Array_StylyColor, Array_StylyQty)














ExitApp




^r::
run % "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " )


ExitApp

Esc::
ExitApp

