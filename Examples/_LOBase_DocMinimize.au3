#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to minimize the document.")

	; Minimize the document.
	_LOBase_DocMinimize($oDoc, True)
	If @error Then _ERROR($oDoc, "Failed to Minimize Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test If document is currently minimized.
	$bReturn = _LOBase_DocMinimize($oDoc)
	If @error Then _ERROR($oDoc, "Failed to query Document status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the document currently minimized? True/False: " & $bReturn & @CRLF & _
			"Press Ok to restore the document to its previous position.")

	; Restore the document to its original size.
	_LOBase_DocMinimize($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to restore Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
