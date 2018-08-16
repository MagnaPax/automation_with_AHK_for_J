driver:= ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver
driver.Get("http://the-automator.com/")
 
;!!!!DO NOT USE!!!  driver.findElement(By.id("search_form_input_homepage")).SendKeys("hello") ; Java bindings of Selenium. !!!!DO NOT USE!!!
 
MsgBox % driver.findElementByID("site-description").Attribute("innerText") ;note case sensitive
MsgBox % driver.executeScript("return document.getElementById('site-description').innerText") ; 이것도 위와 똑같은 작동 하는데 자바 스크립트를 직접적으로 주입(?)한 코드 표현법
MsgBox % driver.executeScript("return document.getElementById('site-description').outerHTML")
 
MsgBox % driver.findElementByID("cat").Attribute("value") ;lowercase value
MsgBox % driver.findElementsByName("cat").item[1].Attribute("outerHTML")
MsgBox % driver.findElementsByName("cat").item[1].Attribute("textContent")
MsgBox % driver.findElementsByName("cat").item[1].Attribute("innerText")
MsgBox % driver.findElementsByName("cat").item[1].Attribute("option value")
MsgBox % driver.findElementByName("s").Attribute("innerTEXT")


Esc::
Exitapp