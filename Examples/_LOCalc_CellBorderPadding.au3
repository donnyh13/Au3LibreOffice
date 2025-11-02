#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $iMicrometers, $iMicrometers2
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B2
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(0.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell's Border padding to 1/4"
	_LOCalc_CellBorderPadding($oCell, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set the Cell's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_CellBorderPadding($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cell's Border padding settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avSettings[0] & " This setting will be either Null or an Integer, depending on whether all four distances are equal or not." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avSettings[3] & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press Ok, and I will demonstrate setting individual border padding settings.")

	; Convert 1/2" to Micrometers
	$iMicrometers2 = _LO_ConvertToMicrometer(0.5)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell's Border padding to, Top and Right, 1/4", Bottom and left, 1/2".
	_LOCalc_CellBorderPadding($oCell, Null, $iMicrometers, $iMicrometers2, $iMicrometers2, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set the Cell's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_CellBorderPadding($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current paragraph Border color settings are as follows: " & @CRLF & _
			"All Padding distance, in Micrometers: " & $avSettings[0] & " This will be Null because the four padding distances aren't equal." & @CRLF & _
			"Top Padding distance, in Micrometers: " & $avSettings[1] & @CRLF & _
			"Bottom Padding distance, in Micrometers: " & $avSettings[2] & @CRLF & _
			"Left Padding distance, in Micrometers: " & $avSettings[3] & @CRLF & @CRLF & _
			"Right Padding distance, in Micrometers: " & $avSettings[4])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
