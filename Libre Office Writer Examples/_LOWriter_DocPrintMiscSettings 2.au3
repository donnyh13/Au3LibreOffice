
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew
	Local $aPrinters
	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now show the currently set Printer.")

	;Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintMiscSettings($oDoc)
	If (@error > 0) Then _ERROR("Error retrieving Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Current Printer Setting", "Your currently set printer name is: " & $avSettings[1] & @CRLF & @CRLF & _
			"I will now modify the setting and show the result.")

	;Retrieve Array of Printers
	$aPrinters = _LOWriter_DocEnumPrintersAlt()
	If (@error > 0) Or Not IsArray($aPrinters) Then _ERROR("Error retrieving Array of Printers. Error:" & @error & " Extended:" & @extended)

	If ($aPrinters[0] <= 1) Then _ERROR("one or no Printers found. Error:" & @error & " Extended:" & @extended)

	;Changes the printer Name setting to another Printer.
	_LOWriter_DocPrintMiscSettings($oDoc, Null, $aPrinters[2])
	If (@error > 0) Then _ERROR("Error setting Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended)

	;Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintMiscSettings($oDoc)
	If (@error > 0) Then _ERROR("Error retrieving Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended)

	;Display the new settings.
	MsgBox($MB_OK, "Current Printer Setting", "Your new set printer name is: " & $avSettingsNew[1] & @CRLF & @CRLF & _
			"I will now return the setting to its original value, and close the document.")

	;Return the setting to its original value by using the previous array of settings I retrieved.
	_LOWriter_DocPrintMiscSettings($oDoc, Null, $avSettings[1])
	If (@error > 0) Then _ERROR("Error restoring Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended)

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



