#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $bResult

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

; Test if the document contains a Sheet called "Sheet1"
$bResult = _LOCalc_DocHasSheetName($oDoc, "Sheet1")
If @error Then _ERROR("Failed to check if a Sheet existed in a Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does this document contain a Sheet named ""Sheet1"" ? True/False: " & $bResult)

; Test if the document contains a Sheet called "FakeSheet"
$bResult = _LOCalc_DocHasSheetName($oDoc, "FakeSheet")
If @error Then _ERROR("Failed to check if a Sheet existed in a Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does this document contain a Sheet named ""FakeSheet"" ? True/False: " & $bResult)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
