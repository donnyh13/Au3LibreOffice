#include <MsgBoxConstants.au3>
#include <Array.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asUndo[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Begin a Undo Action Group record. Name it "AutoIt Insert String"
	_LOWriter_DocUndoActionBegin($oDoc, "AutoIt Insert String")
	If @error Then _ERROR($oDoc, "Failed to begin an Undo Group record. Error:" & @error & " Extended:" & @extended)

	; Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some more text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Some more text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some more text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "One more line of text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available undo action titles.
	$asUndo = _LOWriter_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Here is a list of available Undo Actions. Notice it lists only our Undo Action Group, not any of the other string inserts." & @CRLF & _
			"That is because I have not ended the undo Action Group yet.")

	; Display the available Undo action titles.
	_ArrayDisplay($asUndo)

	; End the Undo Action Record.
	_LOWriter_DocUndoActionEnd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to end an Undo Group record. Error:" & @error & " Extended:" & @extended)

	; Insert some more text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "New text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available undo action titles again.
	$asUndo = _LOWriter_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Here is a list of available Undo Actions. Notice it still lists our Undo Action Group, and also the paragraph and string insert I did after ending the Undo Group record.")

	; Display the available Undo action titles again, if any.
	_ArrayDisplay($asUndo)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
