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

	; Retrieve Cell Range A1-C3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Query the Cell Range for empty cells.
	$aoRanges = _LOCalc_RangeQueryEmpty($oCellRange)
	If @error Then _ERROR($oDoc, "Failed to Query Cell for Empty Cells. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iResults = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now highlight in yellow the cell ranges that are empty in the Cell Range.")

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
	Local $oCell

	; Retrieve Cell A1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell A1 to a Number
	_LOCalc_CellValue($oCell, 10)
	If @error Then _ERROR($oDoc, "Failed to set Cell Number. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A3
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell A3 to a Number
	_LOCalc_CellValue($oCell, 15)
	If @error Then _ERROR($oDoc, "Failed to set Cell Number. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B2
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B2 to a Number
	_LOCalc_CellValue($oCell, 22)
	If @error Then _ERROR($oDoc, "Failed to set Cell Number. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
