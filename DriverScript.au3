#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author: Rana Banerjee

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include <ReusableFunctions.au3>
$i = 1
Do
If StringCompare(_GetData($i, "Execute"), "Y") = 0 Then
   $input2_2 = _GetData($i, "TestCaseName")
   $input2 = """" & @ScriptDir & "\" & $input2_2 & ".au3" & """" & " " & $i
   RunWait(@AutoItExe & " " & $input2)
EndIf
$i = $i + 1
Until _GetData($i, "Iteration") = ""



