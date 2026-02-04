#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTextCursor, $oSelection

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "First Line of text." & @CR & _
			"Second line of text." & @CR & _
			"Third line of text." & @CR & _
			"Fourth Line of Text.")
	If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a TextCursor, I called false for $bCreateAtEnd, the cursor will be created at the beginning of the document.
	$oTextCursor = _LOWriter_DocCreateTextCursor($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Text Cursor right 6 spaces, not selecting the text.
	_LOWriter_CursorMove($oTextCursor, $LOW_TEXTCUR_GO_RIGHT, 6, False)
	If @error Then _ERROR($oDoc, "Failed to move TextCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Text Cursor right 4 spaces, selecting the text.
	_LOWriter_CursorMove($oTextCursor, $LOW_TEXTCUR_GO_RIGHT, 4, True)
	If @error Then _ERROR($oDoc, "Failed to move TextCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to select the text selected by the Text Cursor.")

	; Select the text selected by the Text Cursor.
	_LOWriter_DocSelection($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to set selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to get the current selection.")

	; Retrieve the current selection
	$oSelection = _LOWriter_DocSelection($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current selection. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The currently selected text is: " & _LOWriter_DocGetString($oSelection))

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
