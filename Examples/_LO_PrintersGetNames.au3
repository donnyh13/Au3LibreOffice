#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"

Example()

Func Example()
	Local $iCount
	Local $sDefault, $sPrinters = ""
	Local $asPrinters

	; Minimum Libre version is 4.1, Check Libre Office Version.
	If (_LO_VersionGet(True) < 4.1) Then _ERROR("Current Libre Office version lower than 4.1, this function cannot be used." & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will list your currently available printers.")

	; Retrieve Array of available printers.
	$asPrinters = _LO_PrintersGetNames()
	$iCount = @extended
	If @error Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; If results, display them, else display a message and exit.
	If $iCount > 0 Then
		For $i = 0 To $iCount - 1
			$sPrinters &= $asPrinters[$i] & @CRLF
		Next
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The printers currently available are:" & @CRLF & $sPrinters)

	Else
		_ERROR("No printers found.")
	EndIf

	; Check Libre version for searching default printer.
	If (_LO_VersionGet(True) < 6.3) Then
		_ERROR("Libre Office version is less than 6.3, I cannot list your default printer." & " On Line: " & @ScriptLineNumber)

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I will list your currently default printer next.")
	EndIf

	; Return default printer.
	$sDefault = _LO_PrintersGetNames(True)
	If @error Then _ERROR("Error retrieving Default Printer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($sDefault = "") Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "You do not have a default printer.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "Your default printer is: " & $sDefault)
	EndIf
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
