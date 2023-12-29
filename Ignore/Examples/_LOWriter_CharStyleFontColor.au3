#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oCharStyle
	Local $avCharStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the Character style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a Character style.")
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

	; Set the Character style to "Example" Character style.
	_LOWriter_CharStyleSet($oDoc, $oViewCursor, "Example")
	If @error Then _ERROR("Failed to set the Character style. Error:" & @error & " Extended:" & @extended)

	; Retrieve the "Example" Character Style object.
	$oCharStyle = _LOWriter_CharStyleGetObj($oDoc, "Example")
	If @error Then _ERROR("Failed to retrieve Character style object. Error:" & @error & " Extended:" & @extended)

	; Set "Example" Character style font color to $LOW_COLOR_RED, Transparency to 50%, and Highlight to $LOW_COLOR_GOLD
	_LOWriter_CharStyleFontColor($oCharStyle, $LOW_COLOR_RED, 50, $LOW_COLOR_GOLD)
	If @error Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avCharStyleSettings = _LOWriter_CharStyleFontColor($oCharStyle)
	If @error Then _ERROR("Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Character style's current font color settings are as follows: " & @CRLF & _
			"The current Font color, in Long color format: " & $avCharStyleSettings[0] & " This value is a bit weird because transparency" & _
			" is set to other than 0%" & @CRLF & @CRLF & _
			"Transparency of the font color, in percentage: " & $avCharStyleSettings[1] & @CRLF & _
			"Current Font highlight color, in long color format: " & $avCharStyleSettings[2] & @CRLF & @CRLF & _
			"I will now set transparency to 0%.")

	; Set "Example" Character style transparency to 0
	_LOWriter_CharStyleFontColor($oCharStyle, Null, 0)
	If @error Then _ERROR("Failed to set the Character style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avCharStyleSettings = _LOWriter_CharStyleFontColor($oCharStyle)
	If @error Then _ERROR("Failed to retrieve the Character style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Character style's current font color settings are as follows: " & @CRLF & _
			"The current Font color in Long color format: " & $avCharStyleSettings[0] & @CRLF & _
			"Transparency of the font color, in percentage: " & $avCharStyleSettings[1] & @CRLF & _
			"Current Font highlight color, in long color format: " & $avCharStyleSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
