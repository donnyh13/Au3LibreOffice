#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oFrameStyle, $oViewCursor, $oFrame
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

	; Set Frame style Border Width settings to: $LOW_BORDERWIDTH_MEDIUM on all four sides.
	_LOWriter_FrameStyleBorderWidth($oFrameStyle, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR("Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleBorderWidth($oFrameStyle)
	If @error Then _ERROR("Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame style's current Border Width settings are as follows: " & @CRLF & _
			"The Top Border Width is, in Micrometers, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The Bottom Border Width is, in Micrometers, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The Left Border Width is, in Micrometers, (see UDF constants): " & $avSettings[2] & @CRLF & _
			"The Right Border Width is, in Micrometers, (see UDF constants): " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
