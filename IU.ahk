/*
MTR1E8C7CF3AD
MTR1E6F37342A
MTR1E645ED31F
고객은 취소했는데 N41에 업데이트가 안 되서 물건을 보낸 후 FG에 업데이트가 안 된 아이템은 계속 남을 수 있을텐데

중간에 잘못돼서 ESC 키 누르면 엑셀에서 읽어온 정보들을 저장해서 놓치는 정보가 없어야 됨

CUST PO 가 검색이 안되면 처음부터 다시 시작하지 말고 Input Period 값을 last 365에서 다른 값으로 바꿨다가 다시 last 365로 바꾸면 값이 나올 수도 있다

백오더 버튼 누를 때 PRE ORDER 날짜가 있는 아이템을 백오더 하면 언제 배송할 지 묻지 않는다. 하지만 날짜가 없는 것(인스탁인 아이템)은 묻는다. 있는 것과 없는 것이 섞여 있을 때도 묻는다

페이지에서 물건 전체를 백오더 하면 새로운 BO 페이지가 생기는 게 아니라 그냥 Order Status 가 Back Ordered로 바뀌고 업데이트 하라고 Update 버튼이 생긴다


##############################################
중요!!!!!!!!!!!!!!!!!!!!!!!!!!!

코드 쓸때는 Xpath 를 체크박스로 해야되는것이 아니라 바로 위에 있는 type="checkbox" 가 있는 곳으로 해야된다. 그러면 제대로 작동한다
체크가 되어있으면 -1을 리턴하고 
체크가 안 되어있으면 0을 리턴한다
MsgBox, % driver.FindElementByXPath(Xpath).isSelected()

##############################################


*/



		; ############################################################################################################
		; 다른 오류로 업데이트 못한것과 ORDER STATUS 가 SHIPPED 로 되어있어서 업데이트 못 한 것의 차이를 어떻게 구분하지?
		; ############################################################################################################



; ############################################################################################################
; FG 와 LAS 에 배송된 아이템을 업데이트 하기
; 쓸데없는 값이 정리된 Invoice && C.M Detail Register 엑셀 파일에서 정보를 읽어서 FG 와 LAS 에 배송된 아이템을 업데이트 하기
; 크롬창의 폰트 크기는 반드시 100%가 되어야 된다. 폰트 크기가 바뀌면 Selenium 이 작동이 안됨
; ############################################################################################################

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; 에러 메세지 경고창 안 뜨게 하는 함수
ComObjError(false)


#Include %A_ScriptDir%\lib\

#Include EXCEL.ahk
#Include FG.ahk
#Include LAS.ahk
#Include ChromeGet.ahk
;~ #Include CommWeb.ahk



 ;엑셀에서 읽은 값 저장하는 배열 선언
Str_Total_ExcelInfo := object()

 ;스타일 번호만 저장하는 배열 선언
Array_StyleNo := object()

 ; 스타일 색깔만 저장하는 배열 선언
Array_StylyColor := object()

 ; 스타일 수량만 저장하는 배열 선언
Array_StylyQty := object()

 ; 해당 Cust PO 에 있는 모든 PO 번호 저장하는 배열
Array_PONumber := object()

; 업데이트 안 된 아이템이 있는지 체크하기 위해
Str_Temp := Object()
Str_CheckUnUpdatedItems := object()




global Array_StyleNo, Array_StylyColor





E_driver := new EXCEL
N_driver := new N41
eL_driver := new eLAMBS
F_driver := new FG
LAS_driver := NEW LA


; 크롬 창을 계속 여는 것을 방지하기 위해

IsThereNoOpenedFGPage = 1
IsThereNoOpenedLASPage = 1




/*
##################################################
;실제 사용할 땐 주석 풀어주기
##################################################
driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
driver.AddArgument("disable-infobars") ; Close Message that 'Chrome is being controlled by automated test software'
driver.AddArgument("--start-maximized") ; Maximize Chrome Browser

driver.Get("https://vendoradmin.fashiongo.net/#/home")

driver.ExecuteScript("document.body.style.zoom = '100%';") ; Set the font of Chrome browser to 100%
driver.close() ; closing just one tab of the browser




MsgBox, 262144, Alert, PLEASE OPEN FashionGo AND LASHOWROOM PAGES BY CHROME BROWSER, THEN FOLLOW STEPS BELOW`n`n`n`n1. SET FONT SIZE OF THE PAGES TO 100`% `n`n2. MAXIMIZE THE BROWSER`n`n3. CLOSE THE BROWSER`n`n4. THEN, CLICK OK BUTTON OF THIS WINDOW TO CONTINUE.

*/


Loop{

	; 변수들 초기화 해주기
	Str_Total_ExcelInfo := []
	Array_StyleNo := []
	Array_StylyColor := []
	Array_PONumber := []
	Str_Temp := []
	Str_CheckUnUpdatedItems := []
	Array_StylyQty := []

	ShippingFee =
	PONumber =
	InvoiceNo =
	TrackingNo =
	UnUpdatedItemExists =
	#ofUnUpdatedItems =
	

	



	; 엑셀에서 정보 읽기. 같은 고객의 모든 정보가 다중배열로 반환됨
	Str_Total_ExcelInfo := E_driver.GetInfoFromExcelThenPutThatInAArrayToIU()
	


/*	
	; ##############################################
	; 만약 필요하면 Str_Total_ExcelInfo[6]에 Shipping Fee 추가하는 동작 첨가하기
	; ##############################################
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



	ShippingFee = 
	PONumber := Str_Total_ExcelInfo[1][1]
	InvoiceNo := Str_Total_ExcelInfo[1][2]
	TrackingNo := Str_Total_ExcelInfo[1][3]
	
;	MsgBox, % ShippingFee . "`n`n`n" . PONumber . "`n`n`n" . InvoiceNo . "`n`n`n" . TrackingNo











; FG 오더 업데이트 하기
if(RegExMatch(PONumber, "imU)MTR")){	
	
	
	; FG 크롬 창을 계속 여는 것을 방지하기 위해 
	; IsThereNoOpenedFGPage 값이 1이면 크롬 창 열고 IsThereNoOpenedFGPage 값을 0으로 설정해서 다음 루프가 실행되도 크롬창을 또 열지 않게
	if(IsThereNoOpenedFGPage == 1){
		
		; 일단 이전 주문에서 LAS를 열었을 수도 있으니까 지금 열려있는 크롬창 닫고 시작하기
		WinClose, ahk_class Chrome_WidgetWin_1
		
		URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소

		; 패션고에서 PO Number 얻기 위해 공통으로 처리하는 과정들을 담은 메소드 실행해서 PO Number 검색결과 얻기
		IsItFromNewOrder = 0
		F_driver.ProcessingCommonStepOfOrderProcessing(URL, PONumber, IsItFromNewOrder)
		
		; IsThereNoOpenedFGPage 값은 0으로 설정해 또 다시 열지 않게하고
		; IsThereNoOpenedLASPage 값은 1로 설정해서 다음 오더가 LAS 면 다시 열게끔 하기 왜냐면 IF문이 시작될 때 열린 크롬창은 다 닫았기 때문에
		IsThereNoOpenedFGPage = 0
		IsThereNoOpenedLASPage = 1		
		
	}
	
	

	; 아이템 업데이트 위해 다중배열 Str_Total_ExcelInfo 에 있는
	; 모든 아이템 번호는 Array_StyleNo 배열에 넣고
	; 모든 아이템 색깔은 Array_StylyColor 배열에 넣고
	; 모든 아이템 수량은 Array_StylyQty 배열에 넣기
	Loop % Str_Total_ExcelInfo.Maxindex(){
		
		Array_StyleNo.Insert(Str_Total_ExcelInfo[A_Index][4])
		Array_StylyColor.Insert(Str_Total_ExcelInfo[A_Index][5])
		Array_StylyQty.Insert(Str_Total_ExcelInfo[A_Index][6])
		
	}



/*
	; 스타일 번호와 색깔이 제대로 들어갔는지 확인하기 위해
	Loop % Array_StyleNo.Maxindex(){
		MsgBox % "Element number " . A_Index . " is " . Array_StyleNo[A_Index]
	}

	Loop % Array_StylyColor.Maxindex(){
		MsgBox % "Element number " . A_Index . " is " . Array_StylyColor[A_Index]
	}
	
	MsgBox, Stop!!
*/



	;~ /*
	; Customer PO는 하나만 있는게 아니다. 뒤에 -BO1-BO1 등이 계속 붙을 수 있다
	; 일단 FG 열어서 Customer PO 로 검색했을 때 나오는 모든 PO 번호를 Array_PONumber 배열에 저장한다
	; UpdateItems 메소드는 이 동작에서 얻은 고객의 모든 PO 번호를 사용해서 아이템을 업데이트 한다
	; 크롬창의 폰트 크기는 반드시 100%가 되어야 된다. 폰트 크기가 바뀌면 Selenium 이 작동이 안됨
	;~ Array_PONumber := F_driver.OpenOrderPageForItemUpdate(PONumber, Array_PONumber)
	Array_PONumber := F_driver.OpenOrderPageForItemUpdate(PONumber)



/*
	; Customer PO로 검색된 모든 PO(뒤에 아무것도 안 붙은 오리지널 PO 번호 포함 뒤에 -BO1 붙은 해당되는 모든 PO 제대로 얻었는지 확인하기 위해
	Loop % Array_PONumber.Maxindex(){
		;~ MsgBox % "Element number " . A_Index . " is " . Array_PONumber[A_Index]
		MsgBox % A_Index " PONumber is`n" . Array_PONumber[A_Index]
	}

	MsgBox, pause
*/




	; OpenOrderPageForItemUpdate 메소드를 한 번 실행하고 나오면 driver 인스턴스(객체)값이 소멸되기 때문에 다시 선언해줘야 됨
	F_driver := New FG



	; 아이템 업데이트 하기
	; 지금 못 보내는 아이템들은 백오더 처리하기
	; 배열을 먼저 읽어서 값이 들어있을때만 아래 메소드 호출하기. 아무런 값도 없으면 아예 호출하지 않기(왜냐면 일단 호출하면 
	Loop % Array_PONumber.Maxindex(){

;		MsgBox % A_Index " PONumber is`n" . Array_PONumber[A_Index]
		
		PONumber := Array_PONumber[A_Index]
		F_driver.UpdateItems(Array_StyleNo, Array_StylyColor, Array_StylyQty, ShippingFee, InvoiceNo, TrackingNo, PONumber)

	}



}















; LASHOWROOM 업데이트 하기
if(RegExMatch(PONumber, "imU)OP")){
	
	
	; LAS 크롬 창을 계속 여는 것을 방지하기 위해 
	; IsThereNoOpenedLASPage 값이 1이면 크롬 창 열고 IsThereNoOpenedLASPage 값을 0으로 설정해서 다음 루프가 실행되도 크롬창을 또 열지 않게
	if(IsThereNoOpenedLASPage == 1){
		
		; 일단 이전 주문에서 LAS를 열었을 수도 있으니까 지금 열려있는 크롬창 닫고 시작하기
		WinClose, ahk_class Chrome_WidgetWin_1
		
		URL = https://admin.lashowroom.com/orders_cur_month.php ; 전체 오더 검색창 주소

		; 패션고에서 PO Number 얻기 위해 공통으로 처리하는 과정들을 담은 메소드 실행해서 PO Number 검색결과 얻기
		LAS_driver.LA_ProcessingCommonStepOfOrderProcessing(URL, PONumber)
		
		; IsThereNoOpenedLASPage 값은 0으로 설정해 또 다시 열지 않게하고
		; IsThereNoOpenedFGPage 값은 1로 설정해서 다음 오더가 LAS 면 다시 열게끔 하기 왜냐면 IF문이 시작될 때 열린 크롬창은 다 닫았기 때문에		
		IsThereNoOpenedLASPage = 0
		IsThereNoOpenedFGPage = 1
		
	}
	



	; 아이템 업데이트 위해 다중배열 Str_Total_ExcelInfo 에 있는
	; 모든 아이템 번호는 Array_StyleNo 배열에 넣고
	; 모든 아이템 색깔은 Array_StylyColor 배열에 넣고
	; 모든 아이템 수량은 Array_StylyQty 배열에 넣기
	Loop % Str_Total_ExcelInfo.Maxindex(){
		
		Array_StyleNo.Insert(Str_Total_ExcelInfo[A_Index][4])
		Array_StylyColor.Insert(Str_Total_ExcelInfo[A_Index][5])
		Array_StylyQty.Insert(Str_Total_ExcelInfo[A_Index][6])
		
	}







	; Customer PO는 하나만 있는게 아니다. 뒤에 -A, -B, -C- 등이 계속 붙을 수 있다
	; 일단 LAS 열어서 Customer PO 로 검색했을 때 나오는 모든 PO 번호를 Array_PONumber 배열에 저장한다
	; UpdateItems 메소드는 이 동작에서 얻은 고객의 모든 PO 번호를 사용해서 아이템을 업데이트 한다
	; 크롬창의 폰트 크기는 반드시 100%가 되어야 된다. 폰트 크기가 바뀌면 Selenium 이 작동이 안됨
	Array_PONumber := LAS_driver.OpenOrderPageForItemUpdate(PONumber, Array_PONumber)
	
	

/*   
		; Array_PONumber 에 값이 제대로 들어갔는지 확인하기 위해
		Loop % Array_PONumber.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Array_PONumber[A_Index]
		}
		
		MsgBox PAUSE
*/
	



	; OpenOrderPageForItemUpdate 메소드를 한 번 실행하고 나오면 driver 인스턴스(객체)값이 소멸되기 때문에 다시 선언해줘야 됨
	LAS_driver := New LA



	; 아이템 업데이트 하기
	; 배열을 먼저 읽어서 값이 들어있을때만 아래 메소드 호출하기. 아무런 값도 없으면 아예 호출하지 않기(왜냐면 일단 호출하면 
	Loop % Array_PONumber.Maxindex(){

;		MsgBox % A_Index " PONumber is`n" . Array_PONumber[A_Index]
		
		PONumber := Array_PONumber[A_Index]
		LAS_driver.UpdateItems(Array_StyleNo, Array_StylyColor, Array_StylyQty, ShippingFee, InvoiceNo, TrackingNo, PONumber)

	}










}














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
				
;				MsgBox, 262144, Alert, %UnUpdatedItem_#% %UnUpdatedItem_Color% `nIS NOT UPDATED, PLEASE CHECK IT.
			}
		}

		
		UnUpdatedItemExists := Str_CheckUnUpdatedItems[1][1] ; 업데이트 안 된 아이템이 있는지 확인하기 위해
		#ofUnUpdatedItems := Str_CheckUnUpdatedItems.Maxindex() ; 업데이트 안 된 아이템 숫자
		
		
		; 만약 UnUpdatedItemExists 안에 값이 있으면 업데이트 안 된 아이템이 있다는 뜻
		if(UnUpdatedItemExists != ""){
			
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			
			i = 1
			Loop % Str_CheckUnUpdatedItems.Maxindex(){
				
				UnUpdatedStyleNo := Str_CheckUnUpdatedItems[i][1]
				UnUpdatedStyleColor := Str_CheckUnUpdatedItems[i][2]				
				
				;~ MsgBox % "Customer PO : " PONumber . "n`n`Invoice No : " . InvoiceNo . "`n`n`n`n" . Str_CheckUnUpdatedItems.Maxindex() . " ITEMS ARE NOT UPDATED" . "`n`n" . "Element number " . A_Index . " is " . "`n" . Str_CheckUnUpdatedItems[i][1] . "  " . Str_CheckUnUpdatedItems[i][2]
				MsgBox, 262144, Alert, Customer PO : %PONumber%`n`nInvoice No : %InvoiceNo%`n`n%#ofUnUpdatedItems% ITEMS ARE NOT UPDATED `n`n`n`nElement number %A_Index% is `n`n%UnUpdatedStyleNo%  %UnUpdatedStyleColor%

				i++			
			}
		}

	
	
	
	
	
	
	
	

















;	MsgBox, method out


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





;~ Esc::
;~ Space::
Esc::
ExitApp

