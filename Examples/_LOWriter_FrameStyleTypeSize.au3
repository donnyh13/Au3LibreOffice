
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oFrameStyle
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new FrameStyle named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended)

	;Modify the Frame Style Size settings. Skip Width, set relative width to 50%, Relative to = $LOW_RELATIVE_PARAGRAPH,
	;AutoWidth = False, Skip height, Relative height = 70%, Relative to = $LOW_RELATIVE_PARAGRAPH, AutoHeight = False, Keep Ratio = False
	_LOWriter_FrameStyleTypeSize($oDoc, $oFrameStyle, Null, 50, $LOW_RELATIVE_PARAGRAPH, Null, Null, 70, $LOW_RELATIVE_PARAGRAPH, False, False)
	If (@error > 0) Then _ERROR("Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleTypeSize($oDoc, $oFrameStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame style's size settings are as follows: " & @CRLF & _
			"The Frame style width is, in Micrometers: " & $avSettings[0] & @CRLF & _
			"The frame style relative width percentage is: " & $avSettings[1] & @CRLF & _
			"The width is relative to what? (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"Automatic width? True/False: " & $avSettings[3] & @CRLF & _
			"The Frame style height is, in Micrometers: " & $avSettings[4] & @CRLF & _
			"The frame style relative height percentage is: " & $avSettings[5] & @CRLF & _
			"The height is relative to what? (See UDF Constants): " & $avSettings[6] & @CRLF & _
			"Automatic Height? True/False: " & $avSettings[7] & @CRLF & _
			"Keep Height width Ratio? True/False: " & $avSettings[8])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



