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

	; Create a new Sheet named "First Sheet".
	$oSheet = _LOCalc_SheetAdd($oDoc, "First Sheet")
	If @error Then _ERROR("Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to make the Sheet named, ""First Sheet"" disappear.")

	; Set "First Sheet to be invisible.
	_LOCalc_SheetVisible($oSheet, False)
	If @error Then _ERROR("Failed to set Sheet visibility. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to make the Sheet named, ""First Sheet"" visible again.")

	; Set "First Sheet to be visible.
	_LOCalc_SheetVisible($oSheet, True)
	If @error Then _ERROR("Failed to set Sheet visibility. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
