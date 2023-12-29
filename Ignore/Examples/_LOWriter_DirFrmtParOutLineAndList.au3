#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly." & @CR & "Next Line" & _
			@CR & "Next Line")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Set the paragraph at the current cursor's location Outline and List settings to, Outline level $LOW_OUTLINE_LEVEL_1, Numbering style,"Numbering 123",
	; Include line in paragraph count? True, Line count value = 1.
	_LOWriter_DirFrmtParOutLineAndList($oDoc, $oViewCursor, $LOW_OUTLINE_LEVEL_1, "Numbering 123", True, 1)
	If @error Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtParOutLineAndList($oDoc, $oViewCursor)
	If @error Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Paragraph Outline and List settings are as follows: " & @CRLF & _
			"What Outline level is this Paragraph style set to? (see UDF constants): " & $avSettings[0] & @CRLF & _
			"What Numbering style is used, If any?: " & $avSettings[1] & @CRLF & _
			"Is this paragraph included in line numbering? True/False: " & $avSettings[2] & @CRLF & _
			"Should line counting restart at this paragraph? If so, what number should it start at? (If not, value will be 0): " & $avSettings[3] & @CRLF & @CRLF & _
			"Press ok to remove direct formating.")

	; Remove direct formatting
	_LOWriter_DirFrmtParOutLineAndList($oDoc, $oViewCursor, Default, Default, Default, Default)
	If @error Then _ERROR("Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
