#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

#include <IE.au3>

;Global $oIE = _IECreate("www.google.com")
;Global $oIE = _IECreate("http://localhost:8080/MyHelloWorld")
;ControlClick("Delete Resources", "Review the information","[CLASS:Button;Text:Con&tinue]")
if WinExists("Internet Explorer","&Close program")=1 Then
	  ControlClick ("Internet Explorer","&Close program", "[CLASS:Button;Text:&Close program]")
   endif


#cs
_IELoadWait($oIE)

_ResultReporter ("Keep navigating to the localhost link after every 5 seconds until the page is displayed or timeout of 10 mins",0)
$timer=0
Do
sleep (5000)
$timer=$timer+5
_IENavigate($oIE, _GetData($iteration, "URL"))
until WinExists("Insert title here","")=1 or $timer=240
_ResultReporter ("Wait till the page is fully loaded",0)
_IELoadWait($oIE)
sleep(4000)
#ce