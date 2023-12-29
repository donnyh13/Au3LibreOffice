#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Minimize the document
	_LOCalc_DocMinimize($oDoc, True)
	If @error Then _ERROR("Failed to Minimize Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to bring the document to the front.")

	_LOCalc_DocToFront($oDoc)
	If @error Then _ERROR("Failed to bring document to the front. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
