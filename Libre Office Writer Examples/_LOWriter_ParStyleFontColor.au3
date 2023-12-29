
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $avParStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	;Set "Default Paragraph Style" font color to $LOW_COLOR_RED, Transparency to 50%, and Highlight to $LOW_COLOR_GOLD
	_LOWriter_ParStyleFontColor($oParStyle, $LOW_COLOR_RED, 50, $LOW_COLOR_GOLD)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleFontColor($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current Font Color settings are as follows: " & @CRLF & _
			"The current Font color is, in Long color format: " & $avParStyleSettings[0] & " This value is a bit weird because transparency" & _
			" is set to other than 0%" & @CRLF & @CRLF & _
			"Transparency of the font color, in percentage: " & $avParStyleSettings[1] & @CRLF & _
			"Current Font highlight color, in long color format: " & $avParStyleSettings[2] & @CRLF & @CRLF & _
			"I'll set transparency to 0% now.")

	;Set "Default Paragraph Style" font Transparency to 100%
	_LOWriter_ParStyleFontColor($oParStyle, Null, 0)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleFontColor($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current Font Color settings are as follows: " & @CRLF & _
			"The current Font color is, in Long color format: " & $avParStyleSettings[0] & @CRLF & _
			"Transparency of the font color, in percentage: " & $avParStyleSettings[1] & @CRLF & _
			"Current Font highlight color, in long color format: " & $avParStyleSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



