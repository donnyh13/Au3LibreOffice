#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $sContent

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the top left most cell, 0, 0, or A1.
	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set A1 Cell Value to 55
	_LOCalc_CellValue($oCell, 55)
	If @error Then _ERROR($oDoc, "Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell A1's Text.
	$sContent = _LOCalc_CellString($oCell)
	If @error Then _ERROR($oDoc, "Failed to Retrieve A1 Cell content. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Cell A1 currently contains the following content: " & $sContent)

	; Retrieve the B2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve B2 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set B2 Cell text to "This is Cell B2"
	_LOCalc_CellString($oCell, "This is Cell B2")
	If @error Then _ERROR($oDoc, "Failed to Set B2 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell B2's Text.
	$sContent = _LOCalc_CellString($oCell)
	If @error Then _ERROR($oDoc, "Failed to Retrieve B2 Cell content. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Cell B2 currently contains the following content: " & $sContent)

	; Retrieve the D1 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "D1")
	If @error Then _ERROR($oDoc, "Failed to retrieve D1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set D1 Cell Formula to "=2+2"
	_LOCalc_CellFormula($oCell, "=2+2")
	If @error Then _ERROR($oDoc, "Failed to Set D1 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell D1's Text.
	$sContent = _LOCalc_CellString($oCell)
	If @error Then _ERROR($oDoc, "Failed to Retrieve D1 Cell content. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Cell D1 currently contains the following content: " & $sContent)

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
