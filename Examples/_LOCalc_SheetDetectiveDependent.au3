#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the top left most cell, A1.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set A1 Cell Value to 55
	_LOCalc_CellValue($oCell, 55)
	If @error Then _ERROR($oDoc, "Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve the A2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A2")
	If @error Then _ERROR($oDoc, "Failed to retrieve A2 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set A2 Cell value to 28
	_LOCalc_CellValue($oCell, 28)
	If @error Then _ERROR($oDoc, "Failed to Set A2 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve the B1 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve B1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set B1 Cell Formula to A1 + A2
	_LOCalc_CellFormula($oCell, "=A1+A2")
	If @error Then _ERROR($oDoc, "Failed to Set B1 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve the B2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve B2 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set B2 Cell Formula to A1 - A2
	_LOCalc_CellFormula($oCell, "=A1-A2")
	If @error Then _ERROR($oDoc, "Failed to Set B2 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve the C3 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve C3 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set C3 Cell Formula to SUM(B1; B2)
	_LOCalc_CellFormula($oCell, "=SUM(B1; B2)")
	If @error Then _ERROR($oDoc, "Failed to Set C3 Cell content. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to mark one level of dependents for cell A1.")

	; Retrieve the A1 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Mark one level of dependents for Cell A1
	_LOCalc_SheetDetectiveDependent($oCell)
	If @error Then _ERROR($oDoc, "Failed to mark A1 Cell dependents. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to mark one more level of dependents for cell A1.")

	; Mark one level of dependents for Cell A1
	_LOCalc_SheetDetectiveDependent($oCell)
	If @error Then _ERROR($oDoc, "Failed to mark A1 Cell dependents. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to remove one level of dependent markings for cell A1.")

	; Remove one level of dependent markings for Cell A1
	_LOCalc_SheetDetectiveDependent($oCell, False)
	If @error Then _ERROR($oDoc, "Failed to remove A1 Cell dependents marking. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
