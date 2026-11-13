#include <MsgBoxConstants.au3>

#include "..\LibreOfficeMath.au3"

Example()

Func Example()
	Local $oDoc
	Local $asRedo, $asUndo
	Local $sUndos = "", $sRedos = ""

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOMath_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Math Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOMath_MathFormula($oDoc, "{2 over 2}")
	If @error Then _ERROR($oDoc, "Failed to set Math formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to perform an undo clear, this will work, but will appear to do nothing because undo/redo actions aren't recorded when done using the API." & @CRLF & _
			"If you want this to succeed, please write some text in the formula, and then undo once.")

	; Retrieve an array of available undo action titles.
	$asUndo = _LOMath_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available Redo action titles.
	$asRedo = _LOMath_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $sUndo In $asUndo
		$sUndos &= $sUndo & @CRLF
	Next

	; Display the available Undo action titles.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Undo Actions are:" & @CRLF & $sUndos)

	For $sRedo In $asRedo
		$sRedos &= $sRedo & @CRLF
	Next

	; Display the available Redo action titles.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Redo Actions are:" & @CRLF & $sRedos)

	; Clear the Undo/Redo list.
	_LOMath_DocUndoClear($oDoc)
	If @error Then _ERROR($oDoc, "Failed to clear undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have cleared the Undo/Redo Actions lists. I will retrieve the available Undo and Redo Actions lists again and show that they are now empty.")

	; Retrieve an array of available undo action titles again.
	$asUndo = _LOMath_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available Redo action titles again.
	$asRedo = _LOMath_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sUndos = ""

	For $sUndo In $asUndo
		$sUndos &= $sUndo & @CRLF
	Next

	; Display the available Undo action titles.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Undo Actions are:" & @CRLF & $sUndos)

	$sRedos = ""

	For $sRedo In $asRedo
		$sRedos &= $sRedo & @CRLF
	Next

	; Display the available Redo action titles.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Redo Actions are:" & @CRLF & $sRedos)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOMath_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOMath_DocClose($oDoc, False)
	Exit
EndFunc
