#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oRange
	Local $aavData[3]
	Local $avRowData[3], $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Values.
	$avRowData[0] = 0 ; A1
	$avRowData[1] = "Hello" ; B1
	$avRowData[2] = "Au3" ; C1
	$aavData[0] = $avRowData

	$avRowData[0] = 5 ; A2
	$avRowData[1] = "Goodbye" ; B2
	$avRowData[2] = "123" ; C2
	$aavData[1] = $avRowData

	$avRowData[0] = 1 ; A3
	$avRowData[1] = "Test" ; B3
	$avRowData[2] = "String" ; C3
	$aavData[2] = $avRowData

	; Retrieve Cell range A1 to C3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Data Validity rules for the range A1:A3. Only allow numbers between 1 and 10.
	_LOCalc_RangeValidation($oRange, $LOC_VALIDATION_TYPE_WHOLE, $LOC_VALIDATION_COND_BETWEEN, "1", "10")
	If @error Then _ERROR($oDoc, "Failed to set Range Validation settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range B1 to B3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "B1", "B3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Data Validity rules for the range B1:B3. Only allow Strings that are 5 and more characters long.
	_LOCalc_RangeValidation($oRange, $LOC_VALIDATION_TYPE_TEXT_LEN, $LOC_VALIDATION_COND_GREATER_EQUAL, "5")
	If @error Then _ERROR($oDoc, "Failed to set Range Validation settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range C1 to C3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "C1", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Data Validity rules for the range C1:C3. Only allow Strings that are defined in my list, make the list be shown, and sort it in ascending order.
	_LOCalc_RangeValidation($oRange, $LOC_VALIDATION_TYPE_LIST, $LOC_VALIDATION_COND_EQUAL, '"Au3"; "String"', Null, Null, Null, $LOC_VALIDATION_LIST_SORT_ASCENDING)
	If @error Then _ERROR($oDoc, "Failed to set Range Validation settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to mark any invalid data.")

	; Mark all cells containing invalid data
	_LOCalc_SheetDetectiveInvalidData($oSheet)
	If @error Then _ERROR($oDoc, "Failed to mark cells containing invalid data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current validation settings for Range C1:C3, return will be array with values in order of function parameters.
	$avSettings = _LOCalc_RangeValidation($oRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve Range Validation settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Validation settings for the Range C1:C3 is: " & @CRLF & _
			"The Validation type is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Validation condition is (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"The first condition value is: " & $avSettings[2] & @CRLF & _
			"The second condition value is: " & $avSettings[3] & @CRLF & _
			"The Cell currently set as the reference cell is (whether used or not): " & _LOCalc_RangeGetAddressAsName($avSettings[4]) & @CRLF & _
			"Are blank cells currently ignored (allowed)? True/False: " & $avSettings[5] & @CRLF & _
			"The List visibility type is (See UDF Constants): " & $avSettings[6])

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
