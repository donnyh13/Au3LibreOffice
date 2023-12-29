
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

	;Set "Default Paragraph Style" font position to 75% Superscript, and relative size to 50%.
	_LOWriter_ParStylePosition($oParStyle, Null, 75, Null, Null, 50)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStylePosition($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current position settings are as follows: " & @CRLF & _
			"Is Auto-SuperScript? True/False: " & $avParStyleSettings[0] & @CRLF & _
			"Current SuperScript percentage (If Auto, then it will be 14000): " & $avParStyleSettings[1] & @CRLF & _
			"Is Auto-SubScript? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Current SubScript percentage (If Auto, then it will be -14000): " & $avParStyleSettings[3] & @CRLF & _
			"Relative size percentage: " & $avParStyleSettings[4] & @CRLF & @CRLF & _
			"Press ok and I will set SubScript next.")

	;Set "Default Paragraph Style" font position to 75% Subscript
	_LOWriter_ParStylePosition($oParStyle, Null, Null, Null, 75)
	If (@error > 0) Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStylePosition($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's new position settings are as follows: " & @CRLF & _
			"Is Auto-SuperScript? True/False: " & $avParStyleSettings[0] & @CRLF & _
			"Current SuperScript percentage (If Auto, then it will be 14000): " & $avParStyleSettings[1] & @CRLF & _
			"Is Auto-SubScript? True/False: " & $avParStyleSettings[2] & @CRLF & _
			"Current SubScript percentage (If Auto, then it will be -14000): " & $avParStyleSettings[3] & @CRLF & _
			"Relative size percentage: " & $avParStyleSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



