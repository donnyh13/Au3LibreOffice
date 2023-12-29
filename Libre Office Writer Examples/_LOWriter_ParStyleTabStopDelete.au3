
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oParStyle
	Local $iMicrometers, $iTabStop
	Local $bDeleted

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create a TabStop at 1/4" Tab Stop position.
	$iTabStop = _LOWriter_ParStyleTabStopCreate($oParStyle, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	;Delete the newly created Tab Stop.
	$bDeleted = _LOWriter_ParStyleTabStopDelete($oDoc, $oParStyle, $iTabStop)
	If (@error > 0) Then _ERROR("Failed to delete the Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Was the Tab stop successfully deleted? True/False: " & $bDeleted)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



