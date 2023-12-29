
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oTextCursor
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new Text Cursor.
	$oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oTextCursor, "Some text." & @CR & @CR & "Some different text" & @CR & "Another Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Check if the TextCursor is currently at the end of a word.
	$bReturn = _LOWriter_CursorGetStatus($oTextCursor, $LOW_CURSOR_STAT_IS_END_OF_WORD)
	If (@error > 0) Then _ERROR("Failed to retrieve the Text Cursor Status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the TextCursor at the end of a Word? True/False: " & $bReturn & @CRLF & @CRLF & _
			"I will now move the cursor, and test again.")

	;Move the Cursor to left 3 spaces.
	_LOWriter_CursorMove($oTextCursor, $LOW_TEXTCUR_GO_LEFT, 3)
	If (@error > 0) Then _ERROR("Failed to move cursor. Error:" & @error & " Extended:" & @extended)

	;Check if the TextCursor is currently at the end of a word.
	$bReturn = _LOWriter_CursorGetStatus($oTextCursor, $LOW_CURSOR_STAT_IS_END_OF_WORD)
	If (@error > 0) Then _ERROR("Failed to retrieve the Text Cursor status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the TextCursor at the end of a Word? True/False: " & $bReturn)

	;Check if the TextCursor is currently collapsed, meaning whether or not it has anything selected.
	$bReturn = _LOWriter_CursorGetStatus($oTextCursor, $LOW_CURSOR_STAT_IS_COLLAPSED)
	If (@error > 0) Then _ERROR("Failed to retrieve the Text Cursor status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is nothing selected by the TextCursor? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

