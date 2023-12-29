
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew
	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now show your current print Size settings.")

	;Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintSizeSettings($oDoc)
	If (@error > 0) Then _ERROR("Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Current Settings", "Your current print size settings are as follows: " & @CRLF & @CRLF & _
			"Paper format:— " & $avSettings[0] & @CRLF & _
			"0 =$LOW_PAPER_A3; 1 = $LOW_PAPER_A4; 2 = $LOW_PAPER_A5; 3 = $LOW_PAPER_B4; 4 = $LOW_PAPER_B5; 5 = $LOW_PAPER_LETTER;" & _
			"6 = $LOW_PAPER_LEGAL; 7 = $LOW_PAPER_TABLOID; 8 = $LOW_PAPER_USER_DEFINED" & @CRLF & @CRLF & _
			"Paper Width in Micrometers:— " & $avSettings[1] & @CRLF & _
			"Which is " & _LOWriter_ConvertFromMicrometer($avSettings[1]) & " Inches, and " & _
			_LOWriter_ConvertFromMicrometer(Null, $avSettings[1]) & "Centimeters" & @CRLF & @CRLF & _
			"Paper Height in Micrometers:— " & $avSettings[2] & @CRLF & _
			"Which is " & _LOWriter_ConvertFromMicrometer($avSettings[2]) & " Inches, and " & _
			_LOWriter_ConvertFromMicrometer(Null, $avSettings[2]) & "Centimeters" & @CRLF & @CRLF & _
			"I will now modify the settings and show the result.")

	;Changes the print settings to all false.
	_LOWriter_DocPrintSizeSettings($oDoc, $LOW_PAPER_TABLOID) ;,False,False,False,False)
	If (@error > 0) Then _ERROR("Error setting Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintSizeSettings($oDoc)
	If (@error > 0) Then _ERROR("Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Display the new settings.
	MsgBox($MB_OK, "Current Settings", "Your new print size settings are as follows: " & @CRLF & @CRLF & _
			"Paper format:— " & $avSettingsNew[0] & @CRLF & _
			"0 =$LOW_PAPER_A3; 1 = $LOW_PAPER_A4; 2 = $LOW_PAPER_A5; 3 = $LOW_PAPER_B4; 4 = $LOW_PAPER_B5; 5 = $LOW_PAPER_LETTER;" & _
			"6 = $LOW_PAPER_LEGAL; 7 = $LOW_PAPER_TABLOID; 8 = $LOW_PAPER_USER_DEFINED" & @CRLF & @CRLF & _
			"Paper Width in Micrometers:— " & $avSettingsNew[1] & @CRLF & _
			"Which is " & _LOWriter_ConvertFromMicrometer($avSettingsNew[1]) & " Inches, and " & _
			_LOWriter_ConvertFromMicrometer(Null, $avSettingsNew[1]) & "Centimeters" & @CRLF & @CRLF & _
			"Paper Height in Micrometers:— " & $avSettingsNew[2] & @CRLF & _
			"Which is " & _LOWriter_ConvertFromMicrometer($avSettingsNew[2]) & " Inches, and " & _
			_LOWriter_ConvertFromMicrometer(Null, $avSettingsNew[2]) & "Centimeters" & @CRLF & @CRLF & _
			"I will now modify the settings again and show the result.")

	;Changes the print settings to all false.
	_LOWriter_DocPrintSizeSettings($oDoc, Null, $LOW_PAPER_WIDTH_TABLOID, $LOW_PAPER_HEIGHT_JAP_POSTCARD)
	If (@error > 0) Then _ERROR("Error setting Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintSizeSettings($oDoc)
	If (@error > 0) Then _ERROR("Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Display the new settings.
	MsgBox($MB_OK, "Current Settings", "Your new print size settings are as follows: " & @CRLF & @CRLF & _
			"Paper format:— " & $avSettingsNew[0] & @CRLF & _
			"0 =$LOW_PAPER_A3; 1 = $LOW_PAPER_A4; 2 = $LOW_PAPER_A5; 3 = $LOW_PAPER_B4; 4 = $LOW_PAPER_B5; 5 = $LOW_PAPER_LETTER;" & _
			"6 = $LOW_PAPER_LEGAL; 7 = $LOW_PAPER_TABLOID; 8 = $LOW_PAPER_USER_DEFINED" & @CRLF & @CRLF & _
			"Paper Width in Micrometers:— " & $avSettingsNew[1] & @CRLF & _
			"Which is " & _LOWriter_ConvertFromMicrometer($avSettingsNew[1]) & " Inches, and " & _
			_LOWriter_ConvertFromMicrometer(Null, $avSettingsNew[1]) & "Centimeters" & @CRLF & @CRLF & _
			"Paper Height in Micrometers:— " & $avSettingsNew[2] & @CRLF & _
			"Which is " & _LOWriter_ConvertFromMicrometer($avSettingsNew[2]) & " Inches, and " & _
			_LOWriter_ConvertFromMicrometer(Null, $avSettingsNew[2]) & "Centimeters" & @CRLF & @CRLF & _
			"I will now return the settings to their original values, and close the document.")

	;Restore the original settings
	_LOWriter_DocPrintSizeSettings($oDoc, $avSettings[0], $avSettings[1], $avSettings[2])
	If (@error > 0) Then _ERROR("Error restoring Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

