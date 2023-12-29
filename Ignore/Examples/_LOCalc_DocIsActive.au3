#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Create another New, visible, Blank Libre Office Document.
	$oDoc2 = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	$bReturn = _LOCalc_DocIsActive($oDoc)
	If @error Then _ERROR("Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Document 1 the active document? True/False: " & $bReturn)

	$bReturn = _LOCalc_DocIsActive($oDoc2)
	If @error Then _ERROR("Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Document 2 the active document? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close both documents.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	; Close the second document.
	_LOCalc_DocClose($oDoc2, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
