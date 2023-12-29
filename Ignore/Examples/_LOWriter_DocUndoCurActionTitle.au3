#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $sUndo

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "A word to undo and redo: ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Redo")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Perform one undo action.
	_LOWriter_DocUndo($oDoc)
	If @error Then _ERROR("Failed to perform an undo action. Error:" & @error & " Extended:" & @extended)

	; Retrieve the next available Undo action title.
	$sUndo = _LOWriter_DocUndoCurActionTitle($oDoc)
	If @error Then _ERROR("Failed to retrieve next undo action title. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The next undo action title is: " & $sUndo & " Press ok to perform it.")

	; Perform one undo action.
	_LOWriter_DocUndo($oDoc)
	If @error Then _ERROR("Failed to perform an undo action. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
