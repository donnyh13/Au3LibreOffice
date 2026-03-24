#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oRange
	Local $aavData[3]
	Local $avRowData[3]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
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

	; Mark all cells containing invalid data
	_LOCalc_SheetDetectiveInvalidData($oSheet)
	If @error Then _ERROR($oDoc, "Failed to mark cells containing invalid data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the E5 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "E5")
	If @error Then _ERROR($oDoc, "Failed to retrieve E5 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set E5 Cell Formula to A1-A2 / C5*D3 -- This will result in an error because some of those cells do not have values.
	_LOCalc_CellFormula($oCell, "=A1-A2 / C5*D3")
	If @error Then _ERROR($oDoc, "Failed to Set E5 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the C5 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C5")
	If @error Then _ERROR($oDoc, "Failed to retrieve C5 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set C5 Cell Formula to A1 + A2
	_LOCalc_CellFormula($oCell, "=A1+A2")
	If @error Then _ERROR($oDoc, "Failed to Set C5 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the A7 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A7")
	If @error Then _ERROR($oDoc, "Failed to retrieve A7 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A7 Cell Formula to A7 - C3
	_LOCalc_CellFormula($oCell, "=C5 - C3")
	If @error Then _ERROR($oDoc, "Failed to Set A7 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Mark all levels of precedents for Cell A7
	While _LOCalc_SheetDetectivePrecedent($oCell)
		If @error Then _ERROR($oDoc, "Failed to mark Cell precedents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	WEnd

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have caused most types of Detective Markings to be present, press ok to clear them all from the sheet.")

	; Clear all markings from the Sheet
	_LOCalc_SheetDetectiveClear($oSheet)
	If @error Then _ERROR($oDoc, "Failed to clear Detective markings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
