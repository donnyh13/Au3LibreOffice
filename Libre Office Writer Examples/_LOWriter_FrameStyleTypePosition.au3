
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oFrameStyle
	Local $iMicrometers
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new FrameStyle named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended)

	;Convert 1" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(1)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Modify the Frame Style position settings. Horizontal Alignment = $LOW_ORIENT_HORI_CENTER, Skip Horizontal position,
	;Horizontal relation = $LOW_RELATIVE_PAGE, Mirror = True, Vertical align = $LOW_ORIENT_VERT_NONE, Vertical position = 1",
	;Vertical  relation = $LOW_RELATIVE_PAGE_PRINT, Keep inside = True, Anchor = $LOW_ANCHOR_AT_PAGE
	_LOWriter_FrameStyleTypePosition($oFrameStyle, $LOW_ORIENT_HORI_CENTER, Null, $LOW_RELATIVE_PAGE, True, $LOW_ORIENT_VERT_NONE, $iMicrometers, _
			$LOW_RELATIVE_PAGE_PRINT, True, $LOW_ANCHOR_AT_PAGE)
	If (@error > 0) Then _ERROR("Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleTypePosition($oFrameStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame style's position settings are as follows: " & @CRLF & _
			"The Frame style' Horizontal alignment setting is (see UDF Constants): " & $avSettings[0] & @CRLF & _
			"The frame style's Horizontal position is, in Micrometers: " & $avSettings[1] & @CRLF & _
			"The Frame style' Horizontal relation setting is (see UDF Constants): " & $avSettings[2] & @CRLF & _
			"Mirror Frame position? True/False: " & $avSettings[3] & @CRLF & _
			"The Frame style' Vertical alignment setting is (see UDF Constants): " & $avSettings[4] & @CRLF & _
			"The frame style's Vertical position is, in Micrometers: " & $avSettings[5] & @CRLF & _
			"The Frame style' Vertical relation setting is (see UDF Constants): " & $avSettings[6] & @CRLF & _
			"Keep Frame within Text boundaries? True/False: " & $avSettings[7] & @CRLF & _
			"The Frame Style's anchor position is, (see UDF Constants): " & $avSettings[8])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



