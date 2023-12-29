
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>
#include <Array.au3>

Example()

Func Example()
	Local $oDoc, $oParStyle
	Local $iMicrometers
	Local $aiTabstops

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create a Tab Stop for the demonstration.
	_LOWriter_ParStyleTabStopCreate($oParStyle, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create another Tab Stop for the demonstration.
	_LOWriter_ParStyleTabStopCreate($oParStyle, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	;Convert 1" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(1)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create another Tab Stop for the demonstration.
	_LOWriter_ParStyleTabStopCreate($oParStyle, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	;Retrieve an array of tab stop positions.
	$aiTabstops = _LOWriter_ParStyleTabStopList($oParStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Paragraph style tab stop positions. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The number of Tabstops found is " & @extended & @CRLF & @CRLF & "I will now display the array of tab stop positions.")

	;Display the Array.
	_ArrayDisplay($aiTabstops)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

