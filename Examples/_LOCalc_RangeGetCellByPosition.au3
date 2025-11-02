#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oCellRange

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the top left most cell, 0, 0, or A1.
	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Cell Value to 20
	_LOCalc_CellValue($oCell, 20)
	If @error Then _ERROR($oDoc, "Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the cell located at 1, 1,  or B2 Cell.
	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 1, 1)
	If @error Then _ERROR($oDoc, "Failed to retrieve B2 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set B2 Cell text to "Equals"
	_LOCalc_CellString($oCell, "Equals")
	If @error Then _ERROR($oDoc, "Failed to Set B2 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the cell located at 2, 2, or C3 Cell.
	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 2, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve C3 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set C3 Cell formula to "=A1 * 2
	_LOCalc_CellFormula($oCell, "=A1 * 2")
	If @error Then _ERROR($oDoc, "Failed to Set C3 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now Retrieve Cell Range B4 to D6, and Retrieve the Cell positioned at 2, 2, then set its Background color to Red.")

	; Retrieve Cell Range B4 to D6.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B4", "D6")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the cell located at 2, 2, in the Cell Range.
	$oCell = _LOCalc_RangeGetCellByPosition($oCellRange, 2, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell's Background color to Red.
	_LOCalc_CellBackColor($oCell, $LO_COLOR_RED)
	If @error Then _ERROR($oDoc, "Failed to set Cell's Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
