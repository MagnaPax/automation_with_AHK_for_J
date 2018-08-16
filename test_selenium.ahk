
/*
driver:= ComObjCreate("Selenium.WebDriver") ;Web driver
;~ driver.Start("chrome","http://duckduckgo.com/") ;chrome, firefox, ie, phantomjs, edge
driver.Start("chrome","http://www.daum.net/") ;chrome, firefox, ie, phantomjs, edge
driver.Get("/")

MsgBox
 */
 

driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
;~ driver:= ComObjCreate("Selenium.IEDriver") ;Chrome driver
;~  driver:= ComObjCreate("Selenium.FireFoxDriver") ;Chrome driver
;~ driver.Get("http://duckduckgo.com/")
driver.Get("http://naver.com/")
;~ driver.Get("http://vendoradmin.fashiongo.net/")
*/



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


ExitApp
