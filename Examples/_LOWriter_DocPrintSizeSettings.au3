#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now show your current print Size settings.")

	; Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintSizeSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your current print size settings are as follows: " & @CRLF & @CRLF & _
			"Paper format:— " & $avSettings[0] & @CRLF & _
			"0 =$LOW_PAPER_A3;" & @CRLF & _
			"1 = $LOW_PAPER_A4;" & @CRLF & _
			"2 = $LOW_PAPER_A5;" & @CRLF & _
			"3 = $LOW_PAPER_B4;" & @CRLF & _
			"4 = $LOW_PAPER_B5;" & @CRLF & _
			"5 = $LOW_PAPER_LETTER;" & @CRLF & _
			"6 = $LOW_PAPER_LEGAL;" & @CRLF & _
			"7 = $LOW_PAPER_TABLOID;" & @CRLF & _
			"8 = $LOW_PAPER_USER_DEFINED" & @CRLF & @CRLF & _
			"Paper Width in Micrometers:— " & $avSettings[1] & @CRLF & _
			"Which is " & _LO_ConvertFromMicrometer($avSettings[1]) & " Inches, and " & _
			_LO_ConvertFromMicrometer(Null, $avSettings[1]) & "Centimeters" & @CRLF & @CRLF & _
			"Paper Height in Micrometers:— " & $avSettings[2] & @CRLF & _
			"Which is " & _LO_ConvertFromMicrometer($avSettings[2]) & " Inches, and " & _
			_LO_ConvertFromMicrometer(Null, $avSettings[2]) & "Centimeters" & @CRLF & @CRLF & _
			"I will now modify the settings and show the result.")

	; Changes the print size settings to Tabloid.
	_LOWriter_DocPrintSizeSettings($oDoc, $LOW_PAPER_TABLOID)
	If @error Then _ERROR($oDoc, "Error setting Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintSizeSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Display the new settings.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your new print size settings are as follows: " & @CRLF & @CRLF & _
			"Paper format:— " & $avSettingsNew[0] & @CRLF & _
			"0 =$LOW_PAPER_A3;" & @CRLF & _
			"1 = $LOW_PAPER_A4;" & @CRLF & _
			"2 = $LOW_PAPER_A5;" & @CRLF & _
			"3 = $LOW_PAPER_B4;" & @CRLF & _
			"4 = $LOW_PAPER_B5;" & @CRLF & _
			"5 = $LOW_PAPER_LETTER;" & @CRLF & _
			"6 = $LOW_PAPER_LEGAL;" & @CRLF & _
			"7 = $LOW_PAPER_TABLOID;" & @CRLF & _
			"8 = $LOW_PAPER_USER_DEFINED" & @CRLF & @CRLF & _
			"Paper Width in Micrometers:— " & $avSettingsNew[1] & @CRLF & _
			"Which is " & _LO_ConvertFromMicrometer($avSettingsNew[1]) & " Inches, and " & _
			_LO_ConvertFromMicrometer(Null, $avSettingsNew[1]) & "Centimeters" & @CRLF & @CRLF & _
			"Paper Height in Micrometers:— " & $avSettingsNew[2] & @CRLF & _
			"Which is " & _LO_ConvertFromMicrometer($avSettingsNew[2]) & " Inches, and " & _
			_LO_ConvertFromMicrometer(Null, $avSettingsNew[2]) & "Centimeters" & @CRLF & @CRLF & _
			"I will now modify the settings again and show the result.")

	; Changes the print size settings to Tabloid, but set width to Japanese Postcard.
	_LOWriter_DocPrintSizeSettings($oDoc, Null, $LOW_PAPER_WIDTH_TABLOID, $LOW_PAPER_HEIGHT_JAP_POSTCARD)
	If @error Then _ERROR($oDoc, "Error setting Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintSizeSettings($oDoc)
	If @error Then _ERROR($oDoc, "Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Display the new settings.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your new print size settings are as follows: " & @CRLF & @CRLF & _
			"Paper format:— " & $avSettingsNew[0] & @CRLF & _
			"0 =$LOW_PAPER_A3;" & @CRLF & _
			"1 = $LOW_PAPER_A4;" & @CRLF & _
			"2 = $LOW_PAPER_A5;" & @CRLF & _
			"3 = $LOW_PAPER_B4;" & @CRLF & _
			"4 = $LOW_PAPER_B5;" & @CRLF & _
			"5 = $LOW_PAPER_LETTER;" & @CRLF & _
			"6 = $LOW_PAPER_LEGAL;" & @CRLF & _
			"7 = $LOW_PAPER_TABLOID;" & @CRLF & _
			"8 = $LOW_PAPER_USER_DEFINED" & @CRLF & @CRLF & _
			"Paper Width in Micrometers:— " & $avSettingsNew[1] & @CRLF & _
			"Which is " & _LO_ConvertFromMicrometer($avSettingsNew[1]) & " Inches, and " & _
			_LO_ConvertFromMicrometer(Null, $avSettingsNew[1]) & "Centimeters" & @CRLF & @CRLF & _
			"Paper Height in Micrometers:— " & $avSettingsNew[2] & @CRLF & _
			"Which is " & _LO_ConvertFromMicrometer($avSettingsNew[2]) & " Inches, and " & _
			_LO_ConvertFromMicrometer(Null, $avSettingsNew[2]) & "Centimeters" & @CRLF & @CRLF & _
			"I will now return the settings to their original values, and close the document.")

	; Restore the original settings
	_LOWriter_DocPrintSizeSettings($oDoc, $avSettings[0], $avSettings[1], $avSettings[2])
	If @error Then _ERROR($oDoc, "Error restoring Writer Document Print settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
