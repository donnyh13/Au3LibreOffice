
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

	;Turn Header on.
	_LOWriter_PageStyleHeader($oPageStyle, True)
	If (@error > 0) Then _ERROR("Failed to turn Page Style headers on. Error:" & @error & " Extended:" & @extended)

	;Set Page style Header Background color to $LOW_COLOR_LIME, Background color transparent = False
	_LOWriter_PageStyleHeaderAreaColor($oPageStyle, $LOW_COLOR_LIME, False)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an integer.
	$avPageStyleSettings = _LOWriter_PageStyleHeaderAreaColor($oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Header Background color settings are as follows: " & @CRLF & _
			"The Background color is, in Long Color format: " & $avPageStyleSettings[0] & @CRLF & _
			"Is the background color transparent? True/False: " & $avPageStyleSettings[1])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



