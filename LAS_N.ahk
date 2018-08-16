
;~ /*
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
*/

#Include %A_ScriptDir%\lib\
#Include CNewBrow.AHK


#Include N41.ahk
#Include CommN41.ahk


global driver




; ################################################################################# FG 처리 #################################################################################
; ################################################################################# FG 처리 #################################################################################
; ################################################################################# FG 처리 #################################################################################

;~ /*

driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver

N_driver := new N41

CustomerPO = MTR3200A6A30-BO1-BO1
CustomerPO = MTR20BA0F683E
IsItFromNewOrder = 1

ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) ; ##################### 주문 창 열기 #######################

MsgBox, % "PO # " . CustomerPO . " 이거 열렀나?"
*/




	; FG 오더 처리
	ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile){
	
	
		
		BuyerNotes := ""
		AdditionalInfo := ""
		StaffNotes := ""
		

		
		if(RegExMatch(CustomerPO, "imU)MTR")){
			
			
			; 전체 오더 검색창 주소로 이동하기
			URL = https://vendoradmin.fashiongo.net/#/order/orders ; 전체 오더 검색창 주소
			driver := goToURl_AfterLogIn_IfNeeded(driver, URL) ; 원하는 url로 이동

;~ MsgBox, 전체 오더 검색창 화면으로 이동했음

			; 전체 오더 검색창 주소로 이동한 뒤
			; 검색조건을 PO 번호로 바꾼 뒤 PO 번호로 찾기
			driver := findOrdersByPO#(driver, CustomerPO)
			
;~ MsgBox, 검색 조건을 바꿨음

			; 가장 위에 있는 PO 번호를 새탭으로 열기
			driver := openNewTab_clickMostTopPO#(driver, CustomerPO)
			
;~ MsgBox, 새탭에서 열렸음

			; 현재 페이지의 Order Status 가 New Orders 이거나 Back Ordered 일때 Confirmed Orders 로 바꾸기
			driver := changeNewOrders_To_ConfirmedOrders(driver)
			
;~ MsgBox, Order Status 가 바뀌었음

			
			Arr_FGInfo := getInfoOnFG_And_Return_That(driver, CustomerPO, IsItFromNewOrder, IsItFromExcelFile)

MsgBox, pause 1 

			Arr_BillingAdd := Arr_FGInfo[1].Clone()
			Arr_ShippingAdd := Arr_FGInfo[2].Clone()
			Arr_CC := Arr_FGInfo[3].Clone()
			Arr_Memo := Arr_FGInfo[4].Clone()
			ShippingMethodStatus := Arr_FGInfo[5]

			
			BuyerNotes := Arr_Memo[1]
			AdditionalInfo := Arr_Memo[2]
			StaffNotes := Arr_Memo[3]
			CC# := Arr_CC[2]
			
			

			; 필요 없는 문자가 들어있을 경우를 대비해 메모들 값 정리해주기
			BuyerNotes := Trim(BuyerNotes)
			AdditionalInfo := Trim(AdditionalInfo)
			StaffNotes := Trim(StaffNotes)
			BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")
			AdditionalInfo := RegExReplace(AdditionalInfo, "[^a-zA-Z0-9 ]", "")
			StaffNotes := RegExReplace(StaffNotes, "[^a-zA-Z0-9 ]", "")
	
	
	MsgBox, % Arr_BillingAdd

/* 배열로부터 읽기 두 번째 방법
;~ Array:=[1,3,"ㅋㅋ"]
for index, element in Arr_BillingAdd
{
	MsgBox % "Element number " . index . " is " . element
}
*/


			
			; UPS Ground 값은 3이다. 3이 아니면 
MsgBox, % "ShippingMethodStatus : " . ShippingMethodStatus
			if(ShippingMethodStatus != 3)
			{
				SoundPlay, %A_WinDir%\Media\Ring02.wav
				MsgBox, 262144, UPS STATUS, IT IS NOT UPS GROUND SHIPMENT`n`nOK TO CONTINUE
			}


			; 고객정보 업데이트할지 묻지
			; 메모가 있을때만 창 키워서 표시하기
			if(BuyerNotes || AdditionalInfo || StaffNotes){
					
				SoundPlay, %A_WinDir%\Media\Ring02.wav ; Ring03 이 이상하면 Ring02 써보기
				;~ SoundPlay, %A_WinDir%\Media\Ring03.wav
				MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}
			; 메모 내용이 없으면 간단하게 업데이트 할지만 묻기
			else{

				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}

			; No 눌렀으면 고객정보 업데이트 하지 않기
			IfMsgBox, No
			{					
				; 뉴오더일때만 SO Manager 탭 열기
				; 뉴오더가 아니면 Pick Ticket 뽑다가 디클라인 난 뒤 웹페이지 호출했을 수 있으니 So Mangager 탭 여는게 더 귀찮기 때문					
				if(IsItFromNewOrder){
					N_driver.OpenSOManager() ; SO Manager 탭 열기
				}	
					
				return
	;			Reload
			}


			
			; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
			N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
			
			


			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, Title, Go to SO Manager Tab
			N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
			return
	;		Reload


		}
		
		return
	
	
	} ; ProcessingFGOrder(CustomerPO, F_driver, N_driver, IsItFromNewOrder, IsItFromExcelFile) 메소드 끝

















;~ /*
MsgBox
MsgBox
MsgBox
*/

; ################################################################################# LAS 처리 #################################################################################
; ################################################################################# LAS 처리 #################################################################################
; ################################################################################# LAS 처리 #################################################################################



N_driver := new N41

CustomerPO = OP080951319 ; upsg
CustomerPO = OP080551143 ; usps
CustomerPO = OP080551150 ; 3rd day
CustomerPO = OP080451133 ; delivery
CustomerPO = OP080851265 ; consolidation

ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) ; ##################### 주문 창 열기 #######################

MsgBox, out of method



; LASHOWROOM 오더 처리
; 받은 PO 페이지 열어서 정보 읽고 UPDATE 버튼 누른 뒤 읽은 정보 리턴하기
ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver){
	
	
			driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
			driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
			driver.AddArgument("--start-maximized") ; 창 최대화 하기			
	
			
			; LAS 페이지에서 정보 읽어서 저장하기
			; 메소드가 return [Arr_BillingADD, Arr_ShippingADD, Arr_CC, Arr_Memo] 해서 
			; Arr_FGInfo 배열에는 위 순서대로 값이 저장되어 있음
			Arr_LASInfo := processLAS_which_from_Not_New_Orders(driver, CustomerPO)


			Arr_BillingAdd := Arr_LASInfo[1].Clone()
			Arr_ShippingAdd := Arr_LASInfo[2].Clone()
			Arr_CC := Arr_LASInfo[3].Clone()
			Arr_Memo := Arr_LASInfo[4].Clone()
			ShippingStatus := Arr_LASInfo[5].Clone() ; shipping method 상태 저장 UPSG = 1	ETC = 2		LAS CONSOLIDATION = 3
			
			
			; UPSG 가 아니면 경고 메세지 띄우기
			if(ShippingStatus[1] == "1"){
;				MsgBox, 262144, UPSG, It's UPSG
			}
			else if(ShippingStatus[1] == "2"){
				SoundPlay, %A_WinDir%\Media\Ring02.wav
				MsgBox, 262144, Title, It's neither UPSG nor LAS consolidation
			}
			else{
				SoundPlay, %A_WinDir%\Media\Ring02.wav
				MsgBox, 262144, LAS Consolidation, It's LAS Consolidation
			}
			
			
;			MsgBox, % "ShippingStatus[1] : " . ShippingStatus[1]


			
			BuyerNotes := Arr_Memo[1]
		;	AdditionalInfo := Arr_Memo[2] ; 이 정보는 없음
		;	StaffNotes := Arr_Memo[3] ; 이 정보는 없음
		;	CC# := Arr_CC[2] ; 이 정보는 없음


			; 필요 없는 문자가 들어있을 경우를 대비해 메모값 정리해주기
			BuyerNotes := Trim(BuyerNotes)
			BuyerNotes := RegExReplace(BuyerNotes, "[^a-zA-Z0-9 ]", "")
			
			
			
			/* 배열로부터 읽기 첫 번째 방법
			Loop % Arr_ShippingAdd.Maxindex(){
				MsgBox % "Element number " . A_Index . " is " . Arr_ShippingAdd[A_Index]
			}
			*/			
						


			; 메모가 있을때만 창 키워서 표시하기
			;~ if(BuyerNotes || AdditionalInfo || StaffNotes){
			;~ if BuyerNotes not in None
			if(BuyerNotes)
			{				
				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Memo, `n`n`n`n`n`n`n`n`n`n`n%BuyerNotes%`n`n%AdditionalInfo%`n`n%StaffNotes%`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n`n-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=`n`n`nREADY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}
			else{	

				SoundPlay, %A_WinDir%\Media\Ring06.wav
				MsgBox, 4100, Memo, READY TO UPDATE CUSTOMER INFO`n`n`nIF YOU CLICK Yes, IT'LL OPEN Customer Master AND UPDATE THE CUSTOMER'S INFO ON IT.
			}



			; No 눌렀으면 다시 시작
			IfMsgBox, No
			{
	;			Reload
				return
			}
			
			; CustomerInformationEdit_Tab 에서 정보 업데이트 하기
			N_driver.UpdateInfoOnCustomerInformationEdit_Tab(Arr_ShippingAdd, Arr_CC)
			
			
			SoundPlay, %A_WinDir%\Media\Ring06.wav
			MsgBox, 262144, Title, Go to SO Manager Tab
			N_driver.OpenSOManager() ; SO Manager 탭 열고 끝내기
			
			return
			

	
} ; ProcessingLASOrder(CustomerPO, LASFromAll_driver, N_driver) 메소드 끝








			
			
	
	
	
	
	;~ driver.SwitchToNextWindow ; 새로 연 탭으로 콘트롤 옮김
	driver.Get("http://www.google.com")
	
	
	MsgBox
	
	driver.close() ; closing just one tab of the browser
	
	MsgBox




Exitapp

Esc::
 Exitapp





	
