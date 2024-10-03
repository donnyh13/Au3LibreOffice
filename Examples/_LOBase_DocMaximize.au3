#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "Press Ok to Maximize the document.")

	; Maximize the document.
	_LOBase_DocMaximize($oDoc, True)
	If @error Then _ERROR($oDoc, "Failed to Maximize Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test If document is currently maximized.
	$bReturn = _LOBase_DocMaximize($oDoc)
	If @error Then _ERROR($oDoc, "Failed to query Document status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "Is the document currently maximized? True/False: " & $bReturn & @CRLF & _
			"Press Ok to restore the document to its previous size and position.")

	; Restore the document to its previous size.
	_LOBase_DocMaximize($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to restore Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
