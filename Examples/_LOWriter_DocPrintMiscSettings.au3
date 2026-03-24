#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew
	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now show your current miscellaneous print settings.")

	; Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintMiscSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your current miscellaneous print settings are as follows: " & @CRLF & @CRLF & _
			"Paper Orientation:— " & $avSettings[0] & @CRLF & " 0 = $LOW_PAPER_ORIENT_PORTRAIT, 1 = $LOW_PAPER_ORIENT_LANDSCAPE" & @CRLF & @CRLF & _
			"Printer Name:— " & $avSettings[1] & @CRLF & @CRLF & _
			"Comment Print Mode:— " & $avSettings[2] & @CRLF & " 0 = $LOW_PRINT_NOTES_NONE, 1 = $LOW_PRINT_NOTES_ONLY, " & _
			"2 = $LOW_PRINT_NOTES_END, 3 = $LOW_PRINT_NOTES_NEXT_PAGE" & @CRLF & @CRLF & _
			"Print in Brochure? True/False:— " & $avSettings[3] & @CRLF & @CRLF & _
			"Print Brochure Right to Left? True/False:— " & $avSettings[4] & @CRLF & @CRLF & _
			"Print backwards? True/False:— " & $avSettings[5] & @CRLF & @CRLF & _
			"I will now modify the settings and show the result.")

	; Changes the print settings to Landscape, Skip the printer setting, Print comments Only, Print in brochure,
	; print Brochure in Right to Left mode, and print in reverse.
	_LOWriter_DocPrintMiscSettings($oDoc, $LOW_PAPER_ORIENT_LANDSCAPE, Null, $LOW_PRINT_NOTES_ONLY, True, True, True)
	If @error Then _ERROR($oDoc, "Error setting Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintMiscSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Display the new settings.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your new miscellaneous print settings are as follows: " & @CRLF & @CRLF & _
			"Paper Orientation:— " & $avSettingsNew[0] & @CRLF & " : 0 = $LOW_PAPER_ORIENT_PORTRAIT, 1 = $LOW_PAPER_ORIENT_LANDSCAPE" & @CRLF & @CRLF & _
			"Printer Name:— " & $avSettingsNew[1] & @CRLF & @CRLF & _
			"Comment Print Mode:— " & $avSettingsNew[2] & @CRLF & " 0 = $LOW_PRINT_NOTES_NONE, 1 = $LOW_PRINT_NOTES_ONLY, " & _
			"2 = $LOW_PRINT_NOTES_END, 3 = $LOW_PRINT_NOTES_NEXT_PAGE" & @CRLF & @CRLF & _
			"Print in Brochure? True/False:— " & $avSettingsNew[3] & @CRLF & @CRLF & _
			"Print Brochure Right to Left? True/False:— " & $avSettingsNew[4] & @CRLF & @CRLF & _
			"Print backwards? True/False:— " & $avSettingsNew[5] & @CRLF & @CRLF & _
			"I will now return the settings to their original values, and close the document.")

	; Return the settings to their original values by using the previous array of settings I retrieved.
	_LOWriter_DocPrintMiscSettings($oDoc, $avSettings[0], Null, $avSettings[2], $avSettings[3], $avSettings[4], $avSettings[5])
	If @error Then _ERROR($oDoc, "Error restoring Writer Document Misc Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
