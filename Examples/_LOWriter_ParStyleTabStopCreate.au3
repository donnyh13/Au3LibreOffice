
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oParStyle
	Local $iMicrometers, $iTabStop

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Create a TabStop at 1/4" Tab Stop position, Set the fill character to Asc(~) the Tilde key ASCII Value 126.
	; Set alignment To  $LOW_TAB_ALIGN_DECIMAL, and the decimal character to ASC(.) a period, ASCII value 46.
	$iTabStop = _LOWriter_ParStyleTabStopCreate($oParStyle, $iMicrometers, Asc("~"), $LOW_TAB_ALIGN_DECIMAL, Asc("."))
	If (@error > 0) Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The new Tab stop has the position of " & $iTabStop)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



