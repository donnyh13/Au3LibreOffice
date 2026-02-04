#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text." & @CR & @CR & "Some different text" & @CR & "Another Line.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Cursor to the beginning of the document.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current View cursor position, Return will be the Vertical (Y) coordinate, @Extended is the Horizontal (X) coordinate.
	$iReturn = _LOWriter_DocViewCursorGetPosition($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The ViewCursor is located at the following position:" & @CRLF & _
			"Horizontal, measured in Hundredths of a Millimeter (HMM): " & $iReturn & @CRLF & _
			"Vertical, measured in Hundredths of a Millimeter (HMM): " & @extended & @CRLF & @CRLF & _
			"Press ok, and I will now move the cursor to the end of the document.")

	; Move the Cursor to the beginning of the document.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END)
	If @error Then _ERROR($oDoc, "Failed to move cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the View cursor position again.
	$iReturn = _LOWriter_DocViewCursorGetPosition($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The ViewCursor is now located at the following position:" & @CRLF & _
			"Horizontal, measured in Hundredths of a Millimeter (HMM): " & $iReturn & @CRLF & _
			"Vertical, measured in Hundredths of a Millimeter (HMM): " & @extended)

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
