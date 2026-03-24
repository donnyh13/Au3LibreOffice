#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew
	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now show your current print Include settings.")

	; Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintIncludedSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Print include settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your current print include settings are as follows: " & @CRLF & @CRLF & _
			"Print Graphics? True/False:— " & $avSettings[0] & @CRLF & @CRLF & _
			"Print Controls? True/False:— " & $avSettings[1] & @CRLF & @CRLF & _
			"Print Drawings? True/False:— " & $avSettings[2] & @CRLF & @CRLF & _
			"Print Tables? True/False:— " & $avSettings[3] & @CRLF & @CRLF & _
			"Print Hidden Text? True/False:— " & $avSettings[4] & @CRLF & @CRLF & _
			"I will now modify the settings and show the result.")

	; Changes the print settings to all false.
	_LOWriter_DocPrintIncludedSettings($oDoc, False, False, False, False, False)
	If @error Then _ERROR($oDoc, "Error setting Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintIncludedSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Display the new settings.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your new print include settings are as follows: " & @CRLF & @CRLF & _
			"Print Graphics? True/False:— " & $avSettingsNew[0] & @CRLF & @CRLF & _
			"Print Controls? True/False:— " & $avSettingsNew[1] & @CRLF & @CRLF & _
			"Print Drawings? True/False:— " & $avSettingsNew[2] & @CRLF & @CRLF & _
			"Print Tables? True/False:— " & $avSettingsNew[3] & @CRLF & @CRLF & _
			"Print Hidden Text? True/False:— " & $avSettingsNew[4] & @CRLF & @CRLF & _
			"I will now return the settings to their original values, and close the document.")

	_LOWriter_DocPrintIncludedSettings($oDoc, $avSettings[0], $avSettings[1], $avSettings[2], $avSettings[3], $avSettings[4])
	If @error Then _ERROR($oDoc, "Error restoring Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
