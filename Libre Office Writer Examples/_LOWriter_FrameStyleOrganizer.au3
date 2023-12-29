
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

	;Modify the Frame Style Organizer settings. Change the name to "New Frame Name", set the parent style to "OLE", and auto update to True.
	_LOWriter_FrameStyleOrganizer($oDoc, $oFrameStyle, "New Frame Name", "OLE", True)
	If (@error > 0) Then _ERROR("Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameStyleOrganizer($oDoc, $oFrameStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame Style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame style's organizer settings are as follows: " & @CRLF & _
			"The Frame Style's name is: " & $avSettings[0] & @CRLF & _
			"The name of the parent style of this frame style is: " & $avSettings[1] & @CRLF & _
			"Auto update the Frame style settings when a frame with that Style is modified? True/False: " & $avSettings[2] & @CRLF & _
			"Is this frame style hidden in the User Interface? True/False: " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



