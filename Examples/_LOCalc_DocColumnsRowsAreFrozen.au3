#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check whether there are any frozen panes in the document.
	$bReturn = _LOCalc_DocColumnsRowsAreFrozen($oDoc)
	If @error Then _ERROR($oDoc, "Failed to query Document for frozen panes. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Are there any frozen panes in the document? True/False: " & $bReturn)

	; Freeze the first Row in the Document.
	_LOCalc_DocColumnsRowsFreeze($oDoc, 0, 1)
	If @error Then _ERROR($oDoc, "Failed to freeze Document panes. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check again whether there are any frozen panes in the document.
	$bReturn = _LOCalc_DocColumnsRowsAreFrozen($oDoc)
	If @error Then _ERROR($oDoc, "Failed to query Document for frozen panes. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now are there any frozen panes in the document? True/False: " & $bReturn)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")
	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
