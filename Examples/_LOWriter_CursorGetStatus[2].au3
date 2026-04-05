#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oTextCursor
	Local $bReturn

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Text Cursor.
	$oTextCursor = _LOWriter_TextCursorCreate($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oTextCursor, "Some text." & @CR & @CR & "Some different text" & @CR & "Another Line.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the TextCursor is currently at the end of a word.
	$bReturn = _LOWriter_CursorGetStatus($oTextCursor, $LOW_CURSOR_STAT_IS_END_OF_WORD)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Text Cursor Status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the TextCursor at the end of a Word? True/False: " & $bReturn & @CRLF & @CRLF & _
			"I will now move the cursor, and test again.")

	; Move the Cursor to left 3 spaces.
	_LOWriter_CursorMove($oTextCursor, $LOW_TEXTCUR_GO_LEFT, 3)
	If @error Then _ERROR($oDoc, "Failed to move cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the TextCursor is currently at the end of a word.
	$bReturn = _LOWriter_CursorGetStatus($oTextCursor, $LOW_CURSOR_STAT_IS_END_OF_WORD)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Text Cursor status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the TextCursor at the end of a Word? True/False: " & $bReturn)

	; Check if the TextCursor is currently collapsed, meaning whether or not it has anything selected.
	$bReturn = _LOWriter_CursorGetStatus($oTextCursor, $LOW_CURSOR_STAT_IS_COLLAPSED)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Text Cursor status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is nothing selected by the TextCursor? True/False: " & $bReturn)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
