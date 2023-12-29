
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $iMicrometers, $iMicrometers2
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame into the document at the Viewcursor position, and 3000x3000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Convert 1" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(1)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Modify the Frame wrap type settings. Set wrap type to $LOW_WRAP_MODE_LEFT, Left and Right  Spacing to 1/2", and Top and Bottom spacing to 1"
	_LOWriter_FrameWrap($oFrame, $LOW_WRAP_MODE_LEFT, $iMicrometers, $iMicrometers, $iMicrometers2, $iMicrometers2)
	If (@error > 0) Then _ERROR("Failed to set Frame settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameWrap($oFrame)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame's wrap settings are as follows: " & @CRLF & _
			"The Wrap style is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The spacing between the Left edge of the frame and any text is, in Micrometers,: " & $avSettings[1] & @CRLF & _
			"The spacing between the Right edge of the frame and any text is, in Micrometers,: " & $avSettings[2] & @CRLF & _
			"The spacing between the Top edge of the frame and any text is, in Micrometers,: " & $avSettings[3] & @CRLF & _
			"The spacing between the Bottom edge of the frame and any text is, in Micrometers,: " & $avSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

