#include <MsgBoxConstants.au3>
#include <Array.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asUndo[0], $asRedo[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some more text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Some more text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some more text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "One more line of text")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Undo one action.
	_LOWriter_DocUndo($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Undo the last action. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available undo action titles.
	$asUndo = _LOWriter_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available Redo action titles.
	$asRedo = _LOWriter_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended)

	; Display the available Undo action titles.
	_ArrayDisplay($asUndo)

	; Display the available Undo action titles.
	_ArrayDisplay($asRedo)

	; Clear the Undo/Redo list.
	_LOWriter_DocUndoReset($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Reset undo/redo action titles. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have reset the Undo/Redo Actions lists. I will retrieve the available Undo and Redo Actions lists again and show that they are now empty.")

	; Retrieve an array of available undo action titles again.
	$asUndo = _LOWriter_DocUndoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of undo action titles. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of available Redo action titles again.
	$asRedo = _LOWriter_DocRedoGetAllActionTitles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Redo action titles. Error:" & @error & " Extended:" & @extended)

	; Display the available Undo action titles again, if any.
	_ArrayDisplay($asUndo)

	; Display the available Undo action titles.
	_ArrayDisplay($asRedo)

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
