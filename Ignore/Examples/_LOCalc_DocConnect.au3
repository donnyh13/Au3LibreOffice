#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2
	Local $sDocName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document's name
	$sDocName = _LOCalc_DocGetName($oDoc, False)
	If @error Then _ERROR("Failed to retrieve Calc Document name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created a blank L.O. Calc Doc, I will now Connect to it and use the new Object returned to close it.")

	; Connect to the document.
	$oDoc2 = _LOCalc_DocConnect($sDocName)
	If (@error > 0) Or Not IsObj($oDoc2) Then _ERROR("Failed to Connect to Calc Document. Error:" & @error & " Extended:" & @extended)

	; Close the document, don't save changes.
	_LOCalc_DocClose($oDoc2, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
