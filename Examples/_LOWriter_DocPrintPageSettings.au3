#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings, $avSettingsNew
	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now show your current print Page settings.")

	; Call the function with all optional settings left as Null to retrieve the current settings.
	$avSettings = _LOWriter_DocPrintPageSettings($oDoc)
	If @error Then _ERROR("Error retrieving Writer Document Print page settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Current Settings", "Your current print page settings are as follows: " & @CRLF & @CRLF & _
			"Print in Black only? True/False:— " & $avSettings[0] & @CRLF & @CRLF & _
			"Print Left Pages Only? True/False:— " & $avSettings[1] & @CRLF & @CRLF & _
			"Print Right Pages Only? True/False:— " & $avSettings[2] & @CRLF & @CRLF & _
			"Print the Background? True/False:— " & $avSettings[3] & @CRLF & @CRLF & _
			"Print Empty Pages? True/False:— " & $avSettings[4] & @CRLF & @CRLF & _
			"I will now modify the settings and show the result.")

	; Changes the print settings to all false.
	_LOWriter_DocPrintPageSettings($oDoc, True, False, False, False, False)
	If @error Then _ERROR("Error setting Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	; Now retrieve the settings again.
	$avSettingsNew = _LOWriter_DocPrintPageSettings($oDoc)
	If @error Then _ERROR("Error retrieving Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	; Display the new settings.
	MsgBox($MB_OK, "Current Settings", "Your new print page settings are as follows: " & @CRLF & @CRLF & _
			"Print in Black only? True/False:— " & $avSettingsNew[0] & @CRLF & @CRLF & _
			"Print Left Pages Only? True/False:— " & $avSettingsNew[1] & @CRLF & @CRLF & _
			"Print Right Pages Only? True/False:— " & $avSettingsNew[2] & @CRLF & @CRLF & _
			"Print the Background? True/False:— " & $avSettingsNew[3] & @CRLF & @CRLF & _
			"Print Empty Pages? True/False:— " & $avSettingsNew[4] & @CRLF & @CRLF & _
			"I will now return the settings to their original values, and close the document.")

	; Restore the original print settings.
	_LOWriter_DocPrintPageSettings($oDoc, $avSettings[0], $avSettings[1], $avSettings[2], $avSettings[3], $avSettings[4])
	If @error Then _ERROR("Error restoring Writer Document Print settings. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
