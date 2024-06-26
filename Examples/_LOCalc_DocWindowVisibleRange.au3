#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oRange
	Local $sRange

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the currently visible Range.
	$oRange = _LOCalc_DocWindowVisibleRange($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve currently visible range. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Range's address.
	$sRange = _LOCalc_RangeGetAddressAsName($oRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve range Address. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Currently visible range is: " & $sRange)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
