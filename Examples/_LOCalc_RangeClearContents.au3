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

	; Retrieve the top left most cell, A1.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Cell Value to 20
	_LOCalc_CellValue($oCell, 20)
	If @error Then _ERROR($oDoc, "Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the A2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A2")
	If @error Then _ERROR($oDoc, "Failed to retrieve A2 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A2 Cell Value to .50
	_LOCalc_CellValue($oCell, 0.50)
	If @error Then _ERROR($oDoc, "Failed to Set A2 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the A3 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A3")
	If @error Then _ERROR($oDoc, "Failed to retrieve A3 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A3 Cell text to "String"
	_LOCalc_CellString($oCell, "String")
	If @error Then _ERROR($oDoc, "Failed to Set A3 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the A4 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A4")
	If @error Then _ERROR($oDoc, "Failed to retrieve A4 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A4 Cell formula to "=10 - 5 + 2
	_LOCalc_CellFormula($oCell, "=10 - 5 + 2")
	If @error Then _ERROR($oDoc, "Failed to Set A4 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A4
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A4")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to remove numbers and strings from the cell Range A1:A4. Notice Cell A4 contains a formula, and will remain.")

	; Remove Strings and numbers from the cell range, BitOR the two constants together.
	_LOCalc_RangeClearContents($oCellRange, BitOR($LOC_CELL_FLAG_VALUE, $LOC_CELL_FLAG_STRING))
	If @error Then _ERROR($oDoc, "Failed to clear Cell contents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
