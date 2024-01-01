#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $asRedo

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR("Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the top left most cell, 0, 0, or A1.
	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR("Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set A1 Cell text to "A1"
	_LOCalc_CellString($oCell, "A1")
	If @error Then _ERROR("Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended)

	; Retrieve the C2 Cell.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C2")
	If @error Then _ERROR("Failed to retrieve C2 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set C2 Cell text to "C2"
	_LOCalc_CellString($oCell, "C2")
	If @error Then _ERROR("Failed to Set C2 Cell content. Error:" & @error & " Extended:" & @extended)

	; Perform one undo action.
	_LOCalc_DocUndo($oDoc)
	If @error Then _ERROR("Failed to perform an undo action. Error:" & @error & " Extended:" & @extended)

	; Perform another undo action.
	_LOCalc_DocUndo($oDoc)
	If @error Then _ERROR("Failed to perform an undo action. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available redo action titles.
	$asRedo = _LOCalc_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR("Failed to retrieve array of redo action titles. Error:" & @error & " Extended:" & @extended)

	; Display the available Redo action titles.
	_ArrayDisplay($asRedo)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
