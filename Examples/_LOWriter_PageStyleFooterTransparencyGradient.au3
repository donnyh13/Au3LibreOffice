
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Default Page Style's Object, to modify it's settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	;Turn Footer on.
	_LOWriter_PageStyleFooter($oPageStyle, True)
	If (@error > 0) Then _ERROR("Failed to turn Page Style footers on. Error:" & @error & " Extended:" & @extended)

	;Set Footer Background Color to $LOW_COLOR_RED, Color transparent to False.
	_LOWriter_PageStyleFooterAreaColor($oPageStyle, $LOW_COLOR_RED, False)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Set Page style Footer Transparency Gradient settings to: Gradient Type = $LOW_GRAD_TYPE_ELLIPTICAL, XCenter to 75%, YCenter to 45%, Angle to 180 degrees
	;Border to 16%, Start transparency to 10%, End Transparency to 62%
	_LOWriter_PageStyleFooterTransparencyGradient($oDoc, $oPageStyle, $LOW_GRAD_TYPE_ELLIPTICAL, 75, 45, 180, 16, 10, 62)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleFooterTransparencyGradient($oDoc, $oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Footer Transparency Gradient settings are as follows: " & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avPageStyleSettings[1] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avPageStyleSettings[2] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avPageStyleSettings[3] & @CRLF & _
			"The percentage of area not covered by the transparency is: " & $avPageStyleSettings[4] & @CRLF & _
			"The starting transparency percentage is: " & $avPageStyleSettings[5] & @CRLF & _
			"The ending transparency percentage is: " & $avPageStyleSettings[6])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



