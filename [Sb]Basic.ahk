

/*
driver:= ComObjCreate("Selenium.WebDriver") ;Web driver
driver.Start("firefox","http://duckduckgo.com/") ;chrome, firefox, ie, phantomjs, edge
driver.Get("/")
*/


;~ driver:= ComObjCreate("Selenium.IEDriver") ;IE driver
;~ driver:= ComObjCreate("Selenium.FireFoxDriver") ;FireFox driver

driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
driver.AddArgument("disable-infobars") ;'Chrome이 자동화된 테스트 소프트웨어에 의해 제어되고 있습니다.' 라고 뜨는 경고창 없애기
driver.AddArgument("--start-maximized" ; 창 최대화 하기
driver.addargument("--start-minimized") ; 창 최소화 하기


driver.Get("http://duckduckgo.com/") ; 창 열기


driver.setCapability("ignoreZoomSetting", true) ; 이건 되는지 안 되는지 모르겠음
driver.ExecuteScript("document.body.style.zoom = '100%';") ; 브라우저 폰트 크기를 100%로 설정
driver.executeScript("return document.body.style.zoom = '1.5'") ; 브라우저 폰트 크기를 원래 크기의 1.5배로 설정





MsgBox










/*
;downloaded SeleniumBasic (22 MB)
;(note: did *not* need to separately download Selenium)
;https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
;SeleniumBasic-2.0.9.0.exe
;note: it installed to:
;C:\Users\%username%\AppData\Local\SeleniumBasic

;downloaded ChromeDriver 2.29
;https://sites.google.com/a/chromium.org/chromedriver/downloads
;chromedriver_win32.zip
;note: move new version of chromedriver.exe to install folder:
;e.g. C:\Users\%username%\AppData\Local\SeleniumBasic

;IE and Chrome worked, Mozilla timed out

q::
vBrowser := "ie"
vBrowser := "firefox"
vBrowser := "chrome"

;driver:= ComObjCreate("Selenium.WebDriver")
if (vBrowser = "ie")
	driver:= ComObjCreate("Selenium.IEDriver")
if (vBrowser = "firefox")
	driver:= ComObjCreate("Selenium.FireFoxDriver")
if (vBrowser = "chrome")
	driver:= ComObjCreate("Selenium.CHROMEDriver")

;Start doesn't navigate until the Get
driver.Start(vBrowser,"http://duckduckgo.com/")
driver.Get("/")

Sleep 5000

driver.Get("http://the-automator.com/")
MsgBox % driver.findElementByID("site-description").Attribute("innerText")

MsgBox, % "done"

;reload script closes the browser
;return
driver.AddArgument("disable-infobars")
*/



ExitApp


Esc::
 Exitapp