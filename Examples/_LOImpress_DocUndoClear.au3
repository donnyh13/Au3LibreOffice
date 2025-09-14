#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $asUndo, $asRedo
	Local $sUndos, $sRedos

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; For some reason actions done using COM/UNO commands doesn't register in the Undo stack, so I have to simulate a user action using a Dispatch.
	; Insert a new page.
	_LOImpress_DocExecuteDispatch($oDoc, "uno:InsertPageQuick")
	If @error Then _ERROR($oDoc, "Failed to execute a dispatch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Fixed Date Field.
	_LOImpress_DocExecuteDispatch($oDoc, "uno:InsertDateFieldFix")
	If @error Then _ERROR($oDoc, "Failed to execute a dispatch. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Perform one undo action.
	_LOImpress_DocUndo($oDoc)
	If @error Then _ERROR($oDoc, "Failed to perform an undo action. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available undo action titles.
	$asUndo = _LOImpress_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available Redo action titles.
	$asRedo = _LOImpress_DocRedoGetAllActionTitles($oDoc)
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
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Redo Actions are:" & @CRLF & $sRedos & @CRLF & "Press ok to clear all Undo/Redo actions.")

	; Clear the Undo/Redo list.
	_LOImpress_DocUndoClear($oDoc)
	If @error Then _ERROR($oDoc, "Failed to clear undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have cleared the Undo/Redo Actions lists. I will retrieve the available Undo and Redo Actions lists again and show that they are now empty.")

	; Retrieve an array of available undo action titles again.
	$asUndo = _LOImpress_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available Redo action titles again.
	$asRedo = _LOImpress_DocRedoGetAllActionTitles($oDoc)
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
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
