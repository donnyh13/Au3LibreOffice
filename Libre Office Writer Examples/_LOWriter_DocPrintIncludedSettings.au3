
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew
	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now show your current print Include settings.")

	;Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintIncludedSettings($oDoc)
	If (@error > 0) Then _ERROR("Error retrieving Writer Document Print include settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Current Settings", "Your current print include settings are as follows: " & @CRLF & @CRLF & _
			"Print Graphics? True/False:— " & $avSettings[0] & @CRLF & @CRLF & _
			"Print Controls? True/False:— " & $avSettings[1] & @CRLF & @CRLF & _
			"Print Drawings? True/False:— " & $avSettings[2] & @CRLF & @CRLF & _
			"Print Tables? True/False:— " & $avSettings[3] & @CRLF & @CRLF & _
			"Print Hidden Text? True/False:— " & $avSettings[4] & @CRLF & @CRLF & _
			"I will now modify the settings and show the result.")

	;Changes the print settings to all false.
	_LOWriter_DocPrintIncludedSettings($oDoc, False, False, False, False, False)
	If (@error > 0) Then _ERROR("Error setting Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintIncludedSettings($oDoc)
	If (@error > 0) Then _ERROR("Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Display the new settings.
	MsgBox($MB_OK, "Current Settings", "Your new print include settings are as follows: " & @CRLF & @CRLF & _
			"Print Graphics? True/False:— " & $avSettingsNew[0] & @CRLF & @CRLF & _
			"Print Controls? True/False:— " & $avSettingsNew[1] & @CRLF & @CRLF & _
			"Print Drawings? True/False:— " & $avSettingsNew[2] & @CRLF & @CRLF & _
			"Print Tables? True/False:— " & $avSettingsNew[3] & @CRLF & @CRLF & _
			"Print Hidden Text? True/False:— " & $avSettingsNew[4] & @CRLF & @CRLF & _
			"I will now return the settings to their original values, and close the document.")

	_LOWriter_DocPrintIncludedSettings($oDoc, $avSettings[0], $avSettings[1], $avSettings[2], $avSettings[3], $avSettings[4])
	If (@error > 0) Then _ERROR("Error restoring Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



