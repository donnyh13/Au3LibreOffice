
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "First Line of text" & @CR & _
			"Second line of text." & @CR & _
			"Third line of text." & @CR & _
			"Fourth Line of Text.")
	If (@error > 0) Then _ERROR("Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Return the cursor back to the start.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now move the cursor to the right five characters.")

	;Move the View Cursor right 5 characters, without selecting them.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 5, False)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Next I will move the cursor to the right five more characters, this time I will select them.")

	;Move the View Cursor right 5 characters, and select them.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 5, True)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Next I will move the cursor to the end of the document.")

	;Move the View Cursor to the end of the document, don't select text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END, 1, False)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Next I will move the cursor to the beginning of the line, selecting the text.")

	;Move the View Cursor to the start of the last line, select text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START_OF_LINE, 1, True)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)

	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

