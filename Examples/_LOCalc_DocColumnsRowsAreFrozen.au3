#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

; Check whether there are any frozen panes in the document.
$bReturn = _LOCalc_DocColumnsRowsAreFrozen($oDoc)
	If @error Then _ERROR($oDoc, "Failed to query Document for frozen panes. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Are there any frozen panes in the document? True/False: " & $bReturn)

	; Freeze the first Row in the Document.
	_LOCalc_DocColumnsRowsFreeze($oDoc, 0, 1)
	If @error Then _ERROR($oDoc, "Failed to freeze Document panes. Error:" & @error & " Extended:" & @extended)

; Check again whether there are any frozen panes in the document.
$bReturn = _LOCalc_DocColumnsRowsAreFrozen($oDoc)
	If @error Then _ERROR($oDoc, "Failed to query Document for frozen panes. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now are there any frozen panes in the document? True/False: " & $bReturn)

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
