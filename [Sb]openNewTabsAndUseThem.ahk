; ##########################################################################################################################################
; 새로운 탭 연 뒤 그 탭들 사용하기

/*
SwitchToWindowByTitle(title, timeout)	; 작동됨. 탭의 타이틀을 이용해서 특정 탭으로 이동  예) driver.SwitchToWindowByTitle("Window2")
SwitchToNextWindow(timeout)	; 작동됨
SwitchToPreviousWindow()	; 작동됨
SwitchToParentFrame()	; 작동됨

SwitchToWindowByName(name, timeout)
SwitchToFrame(identifier, timeout)
SwitchToAlert(session, timeout)
*/

; ##########################################################################################################################################



driver := ComObjCreate("Selenium.ChromeDriver")

driver.Get("http://www.google.com")

driver.ExecuteScript("window.open();") ; 새 탭 열기 - 아무것도 없는 빈 새 탭 열기


MsgBox


driver.SwitchToNextWindow ; 새로 연 탭으로 콘트롤 옮김
driver.Get("http://the-automator.com/")
tab_title_#1 := driver.title ; 이렇게 탭 타이틀을 저장해 놓으면 다음에 SwitchToWindowByTitle 명령을 이용해 이 탭을 사용할 수 있음


MsgBox


driver.SwitchToPreviousWindow ; 이전에 열었던 탭으로 되돌아감
driver.Get("http://info.elambs.com/")
tab_title_#2 := driver.title


MsgBox


driver.executeScript("window.open('https://vendoradmin.fashiongo.net')") ; 새 탭 열기 - 지정된 url 열기
tab_title_#3 := driver.title


MsgBox



/*
; 아래 두 개는 작동이 안됨. 타이틀을 저장한다고 다 동작되는게 아닌것 같음
driver.SwitchToWindowByTitle(tab_title_#2)
driver.SwitchToWindowByTitle(tab_title_#3)
*/
driver.SwitchToWindowByTitle(tab_title_#1) ; tab_title_#3 변수에 저장되어있는 탭 타이틀의 탭을 이용하기.(세 번째 탭)
driver.Get("https://admin.lashowroom.com")


MsgBox


driver.executeScript("window.open('http://info.elambs.com/','_target','resizable=yes')") ; open new tab with new destiation ; 아예 새로운 창에서 열기


MsgBox







Exitapp

Esc::
 Exitapp



