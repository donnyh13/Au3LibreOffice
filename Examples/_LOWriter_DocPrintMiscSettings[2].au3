#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew
	Local $aPrinters
	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now show the currently set Printer.")

	; Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintMiscSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your currently set printer name is: " & $avSettings[1] & @CRLF & @CRLF & _
			"I will now modify the setting and show the result.")

	; Retrieve Array of Printers
	$aPrinters = _LO_PrintersGetNamesAlt()
	If (@error > 0) Then _ERROR($oDoc, "Error retrieving Array of Printers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If (@extended <= 1) Then _ERROR($oDoc, "one or no Printers found. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Changes the printer Name setting to another Printer.
	_LOWriter_DocPrintMiscSettings($oDoc, Null, $aPrinters[2])
	If @error Then _ERROR($oDoc, "Error setting Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintMiscSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Display the new settings.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your new set printer name is: " & $avSettingsNew[1] & @CRLF & @CRLF & _
			"I will now return the setting to its original value, and close the document.")

	; Return the setting to its original value by using the previous array of settings I retrieved.
	_LOWriter_DocPrintMiscSettings($oDoc, Null, $avSettings[1])
	If @error Then _ERROR($oDoc, "Error restoring Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
