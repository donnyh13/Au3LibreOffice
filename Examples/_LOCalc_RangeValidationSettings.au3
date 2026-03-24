#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oRange
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to C3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Data Validity rules for the range A1:C3. Only allow numbers between 1 and 10.
	_LOCalc_RangeValidation($oRange, $LOC_VALIDATION_TYPE_WHOLE, $LOC_VALIDATION_COND_BETWEEN, "1", "10")
	If @error Then _ERROR($oDoc, "Failed to set Range Validation settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Validation settings, show a Input message when any of the cells are selected, also set an error message.
	_LOCalc_RangeValidationSettings($oRange, True, "AutoIt Input Title", "Please input a whole number between 1-10, unless you want to see an error message. ", True, _
			$LOC_VALIDATION_ERROR_ALERT_WARNING, "AutoIt Error Title", "You have not entered a whole number inside of the range 1-10. Are you sure you want to proceed?")

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Try entering data in any of the cells in the range of A1:C3, also try entering a decimal value, or it out of the range of 1-10 to see the error message.")

	; Retrieve the current Validation settings for the range. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_RangeValidationSettings($oRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve Range Validation settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Validation settings for the Range A1:C3 is: " & @CRLF & _
			"Is an Input Tooltip shown? True/False: " & $avSettings[0] & @CRLF & _
			"The Tooltip Title is: " & $avSettings[1] & @CRLF & _
			"The Tooltip message is: " & $avSettings[2] & @CRLF & _
			"Is a Error message shown when wrong data is entered? True/False: " & $avSettings[3] & @CRLF & _
			"The type of action that occurs when invalid data is entered is (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The Error message Title is: " & $avSettings[5] & @CRLF & _
			"The Error message is: " & $avSettings[6])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")
	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
