#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Object for the Sheet called "Sheet1".
	$oSheet = _LOCalc_SheetGetObjByName($oDoc, "Sheet1")
	If @error Then _ERROR("Failed to retrieve the Sheet Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Sheet's name is: " & _LOCalc_SheetName($oDoc, $oSheet))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
