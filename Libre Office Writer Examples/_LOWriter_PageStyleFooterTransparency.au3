
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iPageStyleSettings

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

	;Set Page style Footer Transparency settings to 55% transparent
	_LOWriter_PageStyleFooterTransparency($oPageStyle, 55)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an integer.
	$iPageStyleSettings = _LOWriter_PageStyleFooterTransparency($oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Footer Transparency percentage is: " & $iPageStyleSettings)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



