#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "First Line of text" & @CR & _
			"Second line of text." & @CR & _
			"Third line of text." & @CR & _
			"Fourth Line of Text.")
	If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Return the cursor back to the start.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now move the cursor to the right five characters.")

	; Move the View Cursor right 5 characters, without selecting them.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 5, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor to the right five more characters, this time I will select them.")

	; Move the View Cursor right 5 characters, and select them.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 5, True)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor to the end of the document.")

	; Move the View Cursor to the end of the document, don't select text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next I will move the cursor to the beginning of the line, selecting the text.")

	; Move the View Cursor to the start of the last line, select text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START_OF_LINE, 1, True)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
