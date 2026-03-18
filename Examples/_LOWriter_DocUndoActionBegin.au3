#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $sUndos = ""
	Local $asUndo[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Some text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some more text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "Some more text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some more text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "One more line of text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available undo action titles.
	$asUndo = _LOWriter_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Here is a list of available Undo Actions. Notice each action, such as inserting a new line is listed." & @CRLF & _
			"I will reset the Undo and Redo Actions lists, insert some more text, but this time group all the actions together as one Undo action, and then show the Undo Actions list again.")

	For $sUndo In $asUndo
		$sUndos &= $sUndo & @CRLF
	Next

	; Display the available Undo action titles.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Undo Actions are:" & @CRLF & $sUndos)

	; Clear the Undo/Redo list.
	_LOWriter_DocUndoReset($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Reset undo/redo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Begin a Undo Action Group record. Name it "AutoIt Insert String"
	_LOWriter_DocUndoActionBegin($oDoc, "AutoIt Insert String")
	If @error Then _ERROR($oDoc, "Failed to begin an Undo Group record. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some more text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "New text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some more text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "Different line, different text.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; End the Undo Action Record.
	_LOWriter_DocUndoActionEnd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to end an Undo Group record. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of available undo action titles again.
	$asUndo = _LOWriter_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sUndos = ""

	For $sUndo In $asUndo
		$sUndos &= $sUndo & @CRLF
	Next

	; Display the available Undo action titles.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Undo Actions are:" & @CRLF & $sUndos)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
