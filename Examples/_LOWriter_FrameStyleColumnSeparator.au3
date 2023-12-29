#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oFrameStyle, $oViewCursor, $oFrame
	Local $iMicrometers
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a new FrameStyle named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR("Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Frame into the document for demonstration.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	; Set the Frame's style to my created style, "Test Style"
	_LOWriter_FrameStyleSet($oDoc, $oFrame, "Test Style")
	If @error Then _ERROR("Failed to set Frame style. Error:" & @error & " Extended:" & @extended)

	; Modify the Frame Style Column count to 4.
	_LOWriter_FrameStyleColumnSettings($oFrameStyle, 4)
	If @error Then _ERROR("Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended)

	; Convert 1/16" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.0625)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Set the Frame Style Column Separator line settings to: Seperator on (True), Line Style = $LOW_LINE_STYLE_SOLID, Line width to 1/16"
	; Line Color to $LOW_COLOR_RED, Height to 75%, Line Position to $LOW_ALIGN_VERT_MIDDLE
	_LOWriter_FrameStyleColumnSeparator($oFrameStyle, True, $LOW_LINE_STYLE_SOLID, $iMicrometers, $LOW_COLOR_RED, 75, $LOW_ALIGN_VERT_MIDDLE)
	If @error Then _ERROR("Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleColumnSeparator($oFrameStyle)
	If @error Then _ERROR("Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame style's current Column Seperator Line settings are as follows: " & @CRLF & _
			"Is Column seperated by a line? True/False: " & $avSettings[0] & @CRLF & _
			"The Seperator Line style is, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The Seperator Line width is, in Micrometers: " & $avSettings[2] & @CRLF & _
			"The Seperator Line color is, in Long color format: " & $avSettings[3] & @CRLF & _
			"The Seperator Line length percentage is: " & $avSettings[4] & @CRLF & _
			"The Seperator Line position is, (see UDF constants): " & $avSettings[5])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
