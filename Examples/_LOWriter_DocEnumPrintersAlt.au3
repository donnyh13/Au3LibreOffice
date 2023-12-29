
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>
#include <Array.au3>

Example()

Func Example()
	Local $sDefault, $sSearch
	Local $asPrinters

	MsgBox($MB_OK, "", "I will list your currently available printers")

	$asPrinters = _LOWriter_DocEnumPrintersAlt()
	If (@error > 0) Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "There were " & @extended & " results.")

	_ArrayDisplay($asPrinters)

	MsgBox($MB_OK, "", "I will list your currently default printer next.")

	$sDefault = _LOWriter_DocEnumPrintersAlt("", True)
	If (@error > 0) Then _ERROR("Error retrieving Default Printer. Error:" & @error & " Extended:" & @extended)

	If ($sDefault = "") Then
		MsgBox($MB_OK, "", "You do not have a default printer.")
	Else
		MsgBox($MB_OK, "", "Your default printer is: " & $sDefault)
	EndIf

	If (MsgBox($MB_YESNO, "", "We will search for a specific printer next, would you like to enter a phrase to search for?") = $IDYES) Then
		$sSearch = InputBox("", "Enter a search term, if the name is not full and exact, use an asterisk (*), such as ""*PDF*""", "*PDF*")
	Else
		MsgBox($MB_OK, "", "Okay, I will search for ""*PDF*""")
		$sSearch = "*PDF*"
	EndIf

	$asPrinters = _LOWriter_DocEnumPrintersAlt($sSearch)
	If (@error > 0) Then _ERROR("Error retrieving array of Printers. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "There were: " & @extended & " results")

	_ArrayDisplay($asPrinters)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

