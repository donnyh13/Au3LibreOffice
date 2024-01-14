#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Add a new Sheet named "New Sheet" after the first sheet.
	$oSheet = _LOCalc_SheetAdd($oDoc, "New Sheet", 1)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Sheet. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to activate the ""New Sheet""")

	; Activate the "New Sheet"
	_LOCalc_SheetActivate($oDoc, $oSheet)
	If @error Then _ERROR($oDoc, "Failed to Activate a Calc Sheet. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
