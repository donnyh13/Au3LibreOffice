
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers
	Local $avPageStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Default Page Style's Object, to modify it's settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	;Convert 1/16" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.0625)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set Page style Column count to 4.
	_LOWriter_PageStyleColumnSettings($oPageStyle, 4)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Set Page style Column Separator line settings to: Seperator on (True), Line Style = $LOW_LINE_STYLE_SOLID, Line width to 1/16"
	;Line Color to $LOW_COLOR_RED, Height to 75%, Line Position to $LOW_ALIGN_VERT_MIDDLE
	_LOWriter_PageStyleColumnSeparator($oPageStyle, True, $LOW_LINE_STYLE_SOLID, $iMicrometers, $LOW_COLOR_RED, 75, $LOW_ALIGN_VERT_MIDDLE)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleColumnSeparator($oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Column Seperator Line settings are as follows: " & @CRLF & _
			"Is Column seperated by a line? True/False: " & $avPageStyleSettings[0] & @CRLF & _
			"The Seperator Line style is, (see UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
			"The Seperator Line width is, in Micrometers: " & $avPageStyleSettings[2] & @CRLF & _
			"The Seperator Line color is, in Long color format: " & $avPageStyleSettings[3] & @CRLF & _
			"The Seperator Line length percentage is: " & $avPageStyleSettings[4] & @CRLF & _
			"The Seperator Line position is, (see UDF constants): " & $avPageStyleSettings[5])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



