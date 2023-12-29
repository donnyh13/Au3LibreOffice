#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Create a new Sheet named "New Sheet".
	_LOCalc_SheetAdd($oDoc, "New Sheet")
	If @error Then _ERROR("Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended)

	; Create a new Sheet named "First Sheet". This will be placed at the end of the list.
	_LOCalc_SheetAdd($oDoc, "First Sheet")
	If @error Then _ERROR("Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Sheet Object for "Sheet1"
	$oSheet = _LOCalc_SheetGetObjByName($oDoc, "Sheet1")
	If @error Then _ERROR("Failed to Retrieve the Sheet's Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to remove ""Sheet1"".")

	; Remove the Sheet named "Sheet1".
	_LOCalc_SheetRemove($oDoc, $oSheet)
	If @error Then _ERROR("Failed to Remove the Sheet. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
