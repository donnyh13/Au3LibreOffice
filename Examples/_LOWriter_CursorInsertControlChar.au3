#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Some text to demonstrate inserting control characters into a document." & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "This is a $LOW_CON_CHAR_PAR_BREAK --> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control Character, a Paragraph break.
	_LOWriter_CursorInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_PAR_BREAK)
	If @error Then _ERROR($oDoc, "Failed to insert Control character. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_LINE_BREAK --> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control Character, a Line break.
	_LOWriter_CursorInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_LINE_BREAK)
	If @error Then _ERROR($oDoc, "Failed to insert Control character. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_HARD_HYPHEN --> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control Character, a hard hyphen.
	_LOWriter_CursorInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_HARD_HYPHEN)
	If @error Then _ERROR($oDoc, "Failed to insert Control character. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_SOFT_HYPHEN --> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control Character, a Soft Hyphen.
	_LOWriter_CursorInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_SOFT_HYPHEN)
	If @error Then _ERROR($oDoc, "Failed to insert Control character. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_HARD_SPACE --> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control Character, a Hard Space.
	_LOWriter_CursorInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_HARD_SPACE)
	If @error Then _ERROR($oDoc, "Failed to insert Control character. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR & "This is a $LOW_CON_CHAR_APPEND_PAR --> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOWriter_CursorInsertControlChar($oDoc, $oViewCursor, $LOW_CON_CHAR_APPEND_PAR)
	If @error Then _ERROR($oDoc, "Failed to insert Control character. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
