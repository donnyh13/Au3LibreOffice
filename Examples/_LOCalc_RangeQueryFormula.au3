#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $aoRanges[0]
	Local $iResults

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_FillCells($oDoc, $oSheet)

	; Retrieve Cell Range A1-A6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Query Cell Range A1-A6 for Formula Errors.
	$aoRanges = _LOCalc_RangeQueryFormula($oCellRange, $LOC_FORMULA_RESULT_TYPE_ERROR)
	If @error Then _ERROR($oDoc, "Failed to Query Cell for Formula Cells with errors. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iResults = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now highlight in yellow the cell ranges that contain a Formula with an error.")

	; Cycle through the results and set the background color to yellow for each Cell range found
	For $i = 0 To $iResults - 1
		_LOCalc_CellBackColor($aoRanges[$i], $LO_COLOR_YELLOW)
		If @error Then _ERROR($oDoc, "Failed to set Cell Range Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _FillCells(ByRef $oDoc, ByRef $oSheet)
	Local $oCellRange
	Local $aavData[6]
	Local $avRowData[1]

	; Fill my arrays with the desired Formula Values I want in Column A.
	$avRowData[0] = "=*2" ; A1 ; This one contains an error
	$aavData[0] = $avRowData

	$avRowData[0] = "=2+2" ; A2
	$aavData[1] = $avRowData

	$avRowData[0] = "=7 * x" ; A3 ; Another error
	$aavData[2] = $avRowData

	$avRowData[0] = "=A1+A2" ; A4 ; Error because A1 has an error.
	$aavData[3] = $avRowData

	$avRowData[0] = "=10*5" ; A5
	$aavData[4] = $avRowData

	$avRowData[0] = "=A5 + 1" ; A6
	$aavData[5] = $avRowData

	; Retrieve Cell range A1 to A6
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with my Data
	_LOCalc_RangeFormula($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
