#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Excel.au3>
#include <Array.au3>

#include <MsgBoxConstants.au3>

Local $sFilePath1 = @ScriptDir & "\TestData.xls" ;This file should already exist
Local $oExcel = _ExcelBookOpen($sFilePath1,0)

If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist - Shame on you!")
    Exit
EndIf

Local $aArray = _ExcelReadSheetToArray($oExcel)
_ArrayDisplay($aArray, "Array using Default Parameters")
_ExcelBookClose($oExcel)