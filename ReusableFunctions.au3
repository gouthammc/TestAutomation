#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author:  Rana Banerjee

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
; Function definitions *****************************************************************************************************************************

#include <Excel.au3>

;Function to kill process and wait till it gets killed
func _KillProcess($proc)
do
sleep(2000)
ProcessClose ($proc)
until ProcessExists ($proc) = 0
EndFunc

;Function to get data from Excel
func _GetData($row, $colname)
Local $oExcel = _ExcelBookOpen(@ScriptDir & "\TestData.xls",0)
_ExcelSheetActivate($oExcel, "TestData")
If @error = 1 Then
    Return ("Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    Return("File does not exist")
    Exit
EndIf
;Find the Column # having the Column Name
Local $found=False
Local $colNo = 1
do
   if StringCompare($colname, _ExcelReadCell($oExcel,1,$colNo)) = 0 Then
   $found = True
   Elseif  StringCompare("", _ExcelReadCell($oExcel,1,$colNo)) = 0 Then
   _ExcelBookClose($oExcel)
   Return "Column Not Found"
   Exit
   Else
   $colNo=$colNo+1
   EndIf
until $found = True
Local $ret = _ExcelReadCell($oExcel,$row+1,$colNo)
_ExcelBookClose($oExcel)
Return $ret
EndFunc

;Modified fucntion for Waiting for a window
func _WinWaitActiveMod ($WinName, $text, $timeout)
   if WinWait ($WinName,$text,$timeout) <> 0 Then
	  WinActivate ($WinName, $text)
	  _ResultReporter("Window activated : " & $WinName & "Text : " & $text,0)
	  Sleep (1000)
   Else
	  _ResultReporter("Window not found : " & $WinName & "Text : " & $text,2)
	  msgbox (0, "Window Not Found", "Window not found : " & $WinName & "Text : " & $text)
	  Exit
   EndIf
EndFunc

;Function for Result Reporting
;0=DONE
;1=PASS
;2=FAIL
;3=WARNING
func _ResultReporter ($action, $status)
_ExcelWriteCell($oResExcel, $action, $ResultRow, 1)
Select
Case $status=0
   _ExcelWriteCell($oResExcel, "DONE", $ResultRow, 2)
Case $status = 1
   _ExcelWriteCell($oResExcel, "PASS", $ResultRow, 2)
Case $status = 2
   _ExcelWriteCell($oResExcel, "FAIL", $ResultRow, 2)
Case $status = 3
   _ExcelWriteCell($oResExcel, "WARNING", $ResultRow, 2)
EndSelect
_ExcelBookSave($oResExcel)
$ResultRow = $ResultRow + 1
EndFunc

;Tear down function
func _tearDown()
Local $iDiff = TimerDiff($hTimer)
_ResultReporter("Total duration of the execution was : " & Round($iDiff/1000,0) & " seconds", 0)
if _ExcelBookClose ($oResExcel,1,0) <> 1 then
   For $i=1 to 10
   ProcessClose("EXCEL.EXE")
   Next
EndIf
Sleep(2000)
EndFunc

