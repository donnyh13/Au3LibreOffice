#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"

Example()

Func Example()
	Local $sDefault, $sSearch, $sPrinters = ""
	Local $asPrinters

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will list your currently available printers")

	$asPrinters = _LO_PrintersGetNamesAlt()
	If @error Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " results.")

	For $i = 0 To UBound($asPrinters) - 1
		$sPrinters &= $asPrinters[$i] & @CRLF
	Next
	MsgBox($MB_OK + $MB_TOPMOST, Default, "The printers currently available are:" & @CRLF & $sPrinters)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will list your currently default printer next.")

	$sDefault = _LO_PrintersGetNamesAlt("", True)
	If @error Then _ERROR("Error retrieving Default Printer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($sDefault = "") Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "You do not have a default printer.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "Your default printer is: " & $sDefault)
	EndIf

	If (MsgBox($MB_YESNO, "", "We will search for a specific printer next, would you like to enter a phrase to search for?") = $IDYES) Then
		$sSearch = InputBox("", "Enter a search term, if the name is not full and exact, use an asterisk (*), such as ""*PDF*""", "*PDF*")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "Okay, I will search for ""*PDF*""")
		$sSearch = "*PDF*"
	EndIf

	$asPrinters = _LO_PrintersGetNamesAlt($sSearch)
	If @error Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were: " & @extended & " results")

	$sPrinters = ""

	For $i = 0 To UBound($asPrinters) - 1
		$sPrinters &= $asPrinters[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The printers returned from the search are:" & @CRLF & $sPrinters)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
