
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oNumStyle, $oViewCursor
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new NumberingStyle named "Test Style"
	$oNumStyle = _LOWriter_NumStyleCreate($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to create a Numbering Style. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style at the View Cursor to the new style.
	_LOWriter_NumStyleSet($oDoc, $oViewCursor, "Test Style")
	If (@error > 0) Then _ERROR("Failed to Set the Numbering Style. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Line 1" & @LF & "Line 1.1" & @LF & "Line 1.2" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 2.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 2)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Level 2" & @LF & "Line 2.1" & @LF & "Line 2.2" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Set the Numbering Style Level for This Paragraph to 3.
	_LOWriter_NumStyleSetLevel($oDoc, $oViewCursor, 3)
	If (@error > 0) Then _ERROR("Failed to set the Numbering Style level. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Line 3" & @LF & "Line 3.1" & @LF & "Line 3.2")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Modify the Numbering Style Customization settings: Modify Level 2, Numbering format = $LOW_NUM_STYLE_ARABIC, Start at 3, Char Style = "Emphasis",
	;Sub levels = 2, Seperator before = ~ , Seperator after = #, Consecutive Num = False.
	_LOWriter_NumStyleCustomize($oDoc, $oNumStyle, 2, $LOW_NUM_STYLE_ARABIC, 4, "Emphasis", 2, "~", "#", False)
	If (@error > 0) Then _ERROR("Failed to set Numbering Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Numbering Style settings for level 2. Return will be an array in order of function parameters, Return will only have
	;seven elements, because Numbering Style is not set to special, so there will not be a Bullet or Char Decimal value.
	$avSettings = _LOWriter_NumStyleCustomize($oDoc, $oNumStyle, 2)
	If (@error > 0) Then _ERROR("Failed to retrieve Numbering Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Numbering style's current Customization settings for level 2 are as follows: " & @CRLF & _
			"The Number format used is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The Numbering starts at: " & $avSettings[1] & @CRLF & _
			"The Character Style to use is: " & $avSettings[2] & @CRLF & _
			"The number of sub levels to include is: " & $avSettings[3] & @CRLF & _
			"The Seperator before the Numbering symbol is: " & $avSettings[4] & @CRLF & _
			"The Seperator After the Numbering symbol is: " & $avSettings[5] & @CRLF & _
			"Consecutively number levels? True/False: " & $avSettings[6])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

