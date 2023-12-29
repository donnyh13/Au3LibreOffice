
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $iZoom, $iReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current zoom value
	$iReturn = _LOWriter_DocZoom($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve current zoom value. Error:" & @error & " Extended:" & @extended)

	$iZoom = Int($iReturn * .75) ;Set my new zoom value to 75% of the current zoom value.

	;Zoom cannot be less than 20% or greater than 600%, if my value is outside of this, set it to 140%
	If ($iZoom < 20) Or ($iZoom > 600) Then $iZoom = 140

	MsgBox($MB_OK, "", "Your current zoom is: " & $iReturn & "%. I will now set the zoom to: " & $iZoom & "%.")

	;Set the zoom to my new value.
	_LOWriter_DocZoom($oDoc, $iZoom)
	If (@error > 0) Then _ERROR("Failed to set zoom value. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current zoom value again.
	$iReturn = _LOWriter_DocZoom($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve current zoom value. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Your new zoom value is: " & $iReturn & "%.")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
