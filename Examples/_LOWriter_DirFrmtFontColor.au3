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
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor right 13 spaces
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 13)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Select the word "Demonstrate".
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 11, True)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Set the selected text's font color to $LOW_COLOR_RED, Transparency to 50%, and Highlight to $LOW_COLOR_GOLD
	_LOWriter_DirFrmtFontColor($oViewCursor, $LOW_COLOR_RED, 50, $LOW_COLOR_GOLD)
	If @error Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtFontColor($oViewCursor)
	If @error Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The selected text's current font color settings are as follows: " & @CRLF & _
			"The current Font color, in Long color format: " & $avSettings[0] & " This value is a bit weird because transparency" & _
			" is set to other than 0%" & @CRLF & @CRLF & _
			"Transparency of the font color, in percentage: " & $avSettings[1] & @CRLF & @CRLF & _
			"Current Font highlight color, in long color format: " & $avSettings[2] & @CRLF & @CRLF & _
			"I will now set transparency to 0%.")

	; Set the selected text's Font transparency to 0%,
	_LOWriter_DirFrmtFontColor($oViewCursor, Null, 0)
	If @error Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtFontColor($oViewCursor)
	If @error Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The selected text's current font color settings are as follows: " & @CRLF & _
			"The current Font color, in Long color format: " & $avSettings[0] & @CRLF & _
			"Transparency of the font color, in percentage: " & $avSettings[1] & @CRLF & @CRLF & _
			"Current Font highlight color, in long color format: " & $avSettings[2] & @CRLF & @CRLF & _
			"Press ok to remove direct formating.")

	; Remove Direct formatting.
	_LOWriter_DirFrmtFontColor($oViewCursor, Default, Default, Default)
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
