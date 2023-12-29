#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $iCount
	Local $sDefault
	Local $asPrinters

	; Minimum Libre version is 4.1, Check Libre Office Version.
	If (_LOWriter_VersionGet(True) < 4.1) Then _ERROR("Current Libre Office version lower than 4.1, this function cannot be used.")

	MsgBox($MB_OK, "", "I will list your currently available printers.")

	; Retrieve Array of available printers.
	$asPrinters = _LOWriter_DocEnumPrinters()
	$iCount = @extended
	If @error Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended)

	; If results, display them, else display a message and exit.
	If $iCount > 0 Then
		_ArrayDisplay($asPrinters)
	Else
		_ERROR("No printers found.")
	EndIf

	; Check Libre version for searching default printer.
	If (_LOWriter_VersionGet(True) < 6.3) Then
		_ERROR("Libre Office version is less than 6.3, I cannot list your default printer.")
	Else
		MsgBox($MB_OK, "", "I will list your currently default printer next.")
	EndIf

	; Return default printer.
	$sDefault = _LOWriter_DocEnumPrinters(True)
	If @error Then _ERROR("Error retrieving Default Printer. Error:" & @error & " Extended:" & @extended)

	If ($sDefault = "") Then
		MsgBox($MB_OK, "", "You do not have a default printer.")
	Else
		MsgBox($MB_OK, "", "Your default printer is: " & $sDefault)
	EndIf

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
