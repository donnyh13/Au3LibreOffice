
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text." & @CR & @CR & "Some different text" & @CR & "Another Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the Cursor to the beginning of the document.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move cursor. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current View cursor position, Return will be the Vertical (Y) coordinate, @Extended is the Horizontal (X) coordinate.
	$iReturn = _LOWriter_DocViewCursorGetPosition($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor position. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The ViewCursor is located at the following positions:" & @CRLF & _
			"Horizontal, measured in Micrometers: " & $iReturn & @CRLF & _
			"Vertical, measured in Micrometers: " & @extended & @CRLF & @CRLF & _
			"Press ok, and I will now move the cursor to the end of the document.")

	;Move the Cursor to the beginning of the document.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END)
	If (@error > 0) Then _ERROR("Failed to move cursor. Error:" & @error & " Extended:" & @extended)

	;Retrieve the View cursor position again.
	$iReturn = _LOWriter_DocViewCursorGetPosition($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor position. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The ViewCursor's is now located at the following positions:" & @CRLF & _
			"Horizontal, measured in Micrometers: " & $iReturn & @CRLF & _
			"Vertical, measured in Micrometers: " & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

