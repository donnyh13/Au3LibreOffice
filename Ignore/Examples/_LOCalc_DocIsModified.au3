#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Check if the document has been modified since being saved or created.
	$bReturn = _LOCalc_DocIsModified($oDoc)
	If @error Then _ERROR("Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Has the document been modified since being created or saved? True/False: " & $bReturn)

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR("Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the top left most cell, 0, 0, or A1.
	$oCell = _LOCalc_SheetGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR("Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set A1 Cell text to "A1"
	_LOCalc_CellText($oCell, "A1")
	If @error Then _ERROR("Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended)

	; Check if the document has been modified since being saved or created.
	$bReturn = _LOCalc_DocIsModified($oDoc)
	If @error Then _ERROR("Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now has the document been modified since being created or saved? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
