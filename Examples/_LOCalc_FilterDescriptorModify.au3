#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell, $oFilterDesc
	Local $aavData[6]
	Local $avRowData[2], $avSettings[0]
	Local $atFilterFields[2]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A and B.
	$avRowData[0] = 1 ; A1
	$avRowData[1] = 8 ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = 457 ; A2
	$avRowData[1] = 2300 ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = 537 ; A3
	$avRowData[1] = 31 ; B3
	$aavData[2] = $avRowData

	$avRowData[0] = 18 ; A4
	$avRowData[1] = 55 ; B4
	$aavData[3] = $avRowData

	$avRowData[0] = 537     ; A5
	$avRowData[1] = 30 ; B5
	$aavData[4] = $avRowData

	$avRowData[0] = 457     ; A6
	$avRowData[1] = 2300 ; B6
	$aavData[5] = $avRowData

	; Retrieve Cell range A1 to B6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell D5
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "D5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create my first Filter Field, I will insert it directly into my Array.
	; Make this Filter Field apply to Column A (0, because Columns are 0 based internally in LibreOffice Calc.)
	; Set Numeric to True, and my value to 10, Skip String, Condition to "Greater" than my value (10). I don't need to worry about Operator, because this is the first Field in my Array.)
	$atFilterFields[0] = _LOCalc_FilterFieldCreate(0, True, 10, "", $LOC_FILTER_CONDITION_GREATER)
	If @error Then _ERROR($oDoc, "Failed to create a Filter Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create my second Filter Field.
	; Make this Filter Field apply to Column B (1)
	; Set Numeric to True, and my value to 0, Skip String, Condition to "Does Not Contain" my value (0).
	; Set Operator to And, because I want to find values higher than 10 in Column A, as long as the value in Column B does not contain a 0.
	$atFilterFields[1] = _LOCalc_FilterFieldCreate(1, True, 0, "", $LOC_FILTER_CONDITION_DOES_NOT_CONTAIN, $LOC_FILTER_OPERATOR_AND)
	If @error Then _ERROR($oDoc, "Failed to create a Filter Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Filter Descriptor.
	; Use my Filter Fields Array I just created, Set Case Sensitive to False, Skip Duplicates to True, Use Regular Expressions and Headers to False,
	; Copy Output = True, and set Output to Cell D5
	$oFilterDesc = _LOCalc_FilterDescriptorCreate($oCellRange, $atFilterFields, False, True, False, False, True, $oCell)
	If @error Then _ERROR($oDoc, "Failed to create a Filter Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Perform a Filter operation on Range A1 to B6.
	_LOCalc_RangeFilter($oCellRange, $oFilterDesc)
	If @error Then _ERROR($oDoc, "Failed to perform Filter Operation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now modify the Filter Descriptor to not skip duplicates, to look for Column B to contain a 0, and output the results to Cell F7.")

	; Retrieve Cell F7
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "F7")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify my second Filter Field for Column B.
	; Set Condition to "Contains" my value (0).
	_LOCalc_FilterFieldModify($atFilterFields[1], Null, Null, Null, Null, $LOC_FILTER_CONDITION_CONTAINS)
	If @error Then _ERROR($oDoc, "Failed to modify a Filter Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify my Filter Descriptor.
	; Use my Filter Fields Array I just Modified, Set Skip Duplicates to False, and set Output to Cell F7
	_LOCalc_FilterDescriptorModify($oCellRange, $oFilterDesc, $atFilterFields, Null, False, Null, Null, Null, $oCell)
	If @error Then _ERROR($oDoc, "Failed to Modify a Filter Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Perform a Filter operation on Range A1 to B6 again.
	_LOCalc_RangeFilter($oCellRange, $oFilterDesc)
	If @error Then _ERROR($oDoc, "Failed to perform Filter Operation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Filter Descriptor settings. Return will be an array in order of function parameters.
	$avSettings = _LOCalc_FilterDescriptorModify($oCellRange, $oFilterDesc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Filter descriptor settings.. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Filter Descriptor's current settings are as follows: " & @CRLF & _
			"The Filter Fields Array has " & UBound($avSettings[0]) & " elements." & @CRLF & _
			"Is the Filtering Case Sensitive? True/False: " & $avSettings[1] & @CRLF & _
			"Are Duplicate results skipped? True/False: " & $avSettings[2] & @CRLF & _
			"Are Regular Expressions used? True/False: " & $avSettings[3] & @CRLF & _
			"Is the first Row considered a Column Header? True/False: " & $avSettings[4] & @CRLF & _
			"Will the Filtering results be output to another Range? True/False: " & $avSettings[5] & @CRLF & _
			"The Filtering Results will be output to this Cell: " & _LOCalc_RangeGetAddressAsName($avSettings[6]) & @CRLF & _
			"Will the Source and Destination Ranges remain linked for future filtering Operations? True/False: " & $avSettings[7] & @CRLF & @CRLF & _
			"Note: The Output cell always contains a Cell Object, even if Copy output is False, always Check if Copy Results is set to True First." & @CRLF & _
			"Source and Destination can only be linked if Source Range is a Named Database Range.")

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
