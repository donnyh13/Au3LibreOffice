#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet1, $oSheet2

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the presently active Sheet.
	$oSheet1 = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR("Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Add a new Sheet named "New Sheet" after the first sheet.
	$oSheet2 = _LOCalc_SheetAdd($oDoc, "New Sheet", 1)
	If @error Then _ERROR("Failed to Create a new Calc Sheet. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is Sheet1 currently Active? True/False: " & _LOCalc_SheetIsActive($oDoc, $oSheet1) & @CRLF & _
			"Is ""New Sheet"" currently active? True/False: " & _LOCalc_SheetIsActive($oDoc, $oSheet2))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
