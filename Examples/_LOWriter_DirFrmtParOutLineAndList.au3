#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly." & @CR & "Next Line" & _
			@CR & "Next Line")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the paragraph at the current cursor's location Outline and List settings to, Outline level $LOW_PAR_OUTLINE_LEVEL_1, Numbering style,"Numbering 123",
	; Include line in paragraph count? True, Line count value = 1.
	_LOWriter_DirFrmtParOutLineAndList($oDoc, $oViewCursor, $LOW_PAR_OUTLINE_LEVEL_1, "Numbering 123", True, 1)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_DirFrmtParOutLineAndList($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Paragraph Outline and List settings are as follows: " & @CRLF & _
			"What Outline level is this Paragraph style set to? (see UDF constants): " & $avSettings[0] & @CRLF & _
			"What Numbering style is used, If any?: " & $avSettings[1] & @CRLF & _
			"Is this paragraph included in line numbering? True/False: " & $avSettings[2] & @CRLF & _
			"Should line counting restart at this paragraph? If so, what number should it start at? (If not, value will be 0): " & $avSettings[3] & @CRLF & @CRLF & _
			"Press ok to remove direct formatting.")

	; Remove direct formatting
	_LOWriter_DirFrmtParOutLineAndList($oDoc, $oViewCursor, Default, Default, Default, Default)
	If @error Then _ERROR($oDoc, "Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
