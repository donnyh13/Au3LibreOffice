#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oRange
	Local $aavData[3]
	Local $avRowData[3]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values.
	$avRowData[0] = 1 ; A1
	$avRowData[1] = 10 ; B1
	$avRowData[2] = 11 ; C1
	$aavData[0] = $avRowData

	$avRowData[0] = 5 ; A2
	$avRowData[1] = 18 ; B2
	$avRowData[2] = 8 ; C2
	$aavData[1] = $avRowData

	$avRowData[0] = 0 ; A3
	$avRowData[1] = 3.2 ; B3
	$avRowData[2] = 22 ; C3
	$aavData[2] = $avRowData

	; Retrieve Cell range A1 to C3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeNumbers($oRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to B3
	$oRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Data Validity rules for the range A1:B3. Only allow numbers between 1 and 10.
	_LOCalc_RangeValidation($oRange, $LOC_VALIDATION_TYPE_WHOLE, $LOC_VALIDATION_COND_BETWEEN, "1", "10")
	If @error Then _ERROR($oDoc, "Failed to set Range Validation settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have set validation rules for the Cell Range A1:B3 to only allow Whole numbers between 1 and 10." & @CRLF & _
			"Press Ok to mark any invalid data in the Sheet.")

	; Mark all cells containing invalid data
	_LOCalc_SheetDetectiveInvalidData($oSheet)
	If @error Then _ERROR($oDoc, "Failed to mark cells containing invalid data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
