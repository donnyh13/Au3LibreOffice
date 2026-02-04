#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oFrameStyle
	Local $iHMM, $iHMM2
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Frame Style named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(1, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame Style wrap type settings. Set wrap type to $LOW_WRAP_MODE_LEFT, Left and Right  Spacing to 1/2", and Top and Bottom spacing to 1"
	_LOWriter_FrameStyleWrap($oFrameStyle, $LOW_WRAP_MODE_LEFT, $iHMM, $iHMM, $iHMM2, $iHMM2)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleWrap($oFrameStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame style's wrap settings are as follows: " & @CRLF & _
			"The Wrap style is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The spacing between the Left edge of the frame and any text is, in Hundredths of a Millimeter (HMM),: " & $avSettings[1] & @CRLF & _
			"The spacing between the Right edge of the frame and any text is, in Hundredths of a Millimeter (HMM),: " & $avSettings[2] & @CRLF & _
			"The spacing between the Top edge of the frame and any text is, in Hundredths of a Millimeter (HMM),: " & $avSettings[3] & @CRLF & _
			"The spacing between the Bottom edge of the frame and any text is, in Hundredths of a Millimeter (HMM),: " & $avSettings[4])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
