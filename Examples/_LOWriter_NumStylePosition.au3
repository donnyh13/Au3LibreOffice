
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oNumStyle, $oViewCursor
	Local $iMicrometers, $iMicrometers2, $iMicrometers3
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
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Line 1" & @LF & "Line 1.1" & @LF & "Line 1.2" & @CR & _
			"Line 2" & @LF & "Line 2.1" & @LF & "Line 2.2" & @CR & "Line 3" & @LF & "Line 3.1" & @LF & "Line 3.2")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Convert 3/4" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(.75)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Convert 1" to Micrometers
	$iMicrometers3 = _LOWriter_ConvertToMicrometer(1)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Modify the Numbering Style position settings: Modify all Levels (0), Aligned at = 1/2", Numbering Align = $LOW_ORIENT_HORI_CENTER,
	;Followed by = $LOW_FOLLOW_BY_TABSTOP, TabStop = 3/4", Indent = 1"
	_LOWriter_NumStylePosition($oDoc, $oNumStyle, 0, $iMicrometers, $LOW_ORIENT_HORI_CENTER, $LOW_FOLLOW_BY_TABSTOP, $iMicrometers2, $iMicrometers3)
	If (@error > 0) Then _ERROR("Failed to set Numbering Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Numbering Style settings for level 2. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_NumStylePosition($oDoc, $oNumStyle, 2)
	If (@error > 0) Then _ERROR("Failed to retrieve Numbering Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Numbering style's current Position settings for level 2 are as follows: " & @CRLF & _
			"The First line indent is, in Micrometers: " & $avSettings[0] & @CRLF & _
			"The Numbering symbols are aligned to, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The Numbering Symbol is followed by, (see UDF Constants): " & $avSettings[2] & @CRLF & _
			"The Tab Stop size following the Numbering symbols is, in Micrometers: " & $avSettings[3] & @CRLF & _
			"The indent from the left page margin to the Numbering symbols is, in micrometers: " & $avSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

