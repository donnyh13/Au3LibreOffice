#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet
	Local $sName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the currently Active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR("Failed to Retrieve the currently Active Sheet's Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Active Sheet's Name.
	$sName = _LOCalc_SheetName($oDoc, $oSheet)
	If @error Then _ERROR("Failed to Retrieve the currently Active Sheet's Name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Currently Active Sheet's name is: " & $sName)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
