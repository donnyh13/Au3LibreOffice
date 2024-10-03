#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the top left most cell, A1.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A1 Cell Value to 55
	_LOCalc_CellValue($oCell, 55)
	If @error Then _ERROR($oDoc, "Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the A2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A2")
	If @error Then _ERROR($oDoc, "Failed to retrieve A2 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set A2 Cell value to 28
	_LOCalc_CellValue($oCell, 28)
	If @error Then _ERROR($oDoc, "Failed to Set A2 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the B1 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve B1 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set B1 Cell Formula to A1 + A2
	_LOCalc_CellFormula($oCell, "=A1+A2")
	If @error Then _ERROR($oDoc, "Failed to Set B1 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the B2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve B2 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set B2 Cell Formula to A1 - A2
	_LOCalc_CellFormula($oCell, "=A1-A2")
	If @error Then _ERROR($oDoc, "Failed to Set B2 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the C3 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve C3 Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set C3 Cell Formula to SUM(B1; B2)
	_LOCalc_CellFormula($oCell, "=SUM(B1; B2)")
	If @error Then _ERROR($oDoc, "Failed to Set C3 Cell content. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to mark one level of precedents for cell C3.")

	; Mark one level of precedents for Cell C3
	_LOCalc_SheetDetectivePrecedent($oCell)
	If @error Then _ERROR($oDoc, "Failed to mark C3 Cell precedents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to mark one more level of precedents for cell C3.")

	; Mark one level of precedents for Cell C3
	_LOCalc_SheetDetectivePrecedent($oCell)
	If @error Then _ERROR($oDoc, "Failed to mark C3 Cell precedents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to remove one level of precedent markings for cell C3.")

	; Mark one level of precedents for Cell C3
	_LOCalc_SheetDetectivePrecedent($oCell, False)
	If @error Then _ERROR($oDoc, "Failed to mark C3 Cell precedents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
