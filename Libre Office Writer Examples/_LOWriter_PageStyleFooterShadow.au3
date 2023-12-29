
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

	;Convert 1/8" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.125)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Turn Footer on.
	_LOWriter_PageStyleFooter($oPageStyle, True)
	If (@error > 0) Then _ERROR("Failed to turn Page Style footers on. Error:" & @error & " Extended:" & @extended)

	;Set Page style Footer Shadow settings to: Width = 1/8", Color = $LOW_COLOR_RED, Transparent = False, Location = $LOW_SHADOW_TOP_LEFT
	_LOWriter_PageStyleFooterShadow($oPageStyle, $iMicrometers, $LOW_COLOR_RED, False, $LOW_SHADOW_TOP_LEFT)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleFooterShadow($oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Footer Shadow settings are as follows: " & @CRLF & _
			"The shadow width is, is Micrometers: " & $avPageStyleSettings[0] & @CRLF & _
			"The Shdaow color is, in Long Color format: " & $avPageStyleSettings[1] & @CRLF & _
			"Is the Color transparent? True/False: " & $avPageStyleSettings[2] & @CRLF & _
			"The Shadow location is, (see UDF Constants): " & $avPageStyleSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



