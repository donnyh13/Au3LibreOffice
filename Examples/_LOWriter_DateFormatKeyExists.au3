
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $iFormatKey
	Local $bExists

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new DateFormatKey H:MM:SS (Hour, Minute, Second, AM/PM, Month Day Name(Day) Year
	$iFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "H:MM:SS AM/PM -- M/NNN(D)/YYYY")
	If (@error > 0) Then _ERROR("Failed to create a Date/Time Format Key. Error:" & @error & " Extended:" & @extended)

	;Check if the new key exists.
	$bExists = _LOWriter_DateFormatKeyExists($oDoc, $iFormatKey)
	If (@error > 0) Then _ERROR("Failed to search for a Date/Time Format Key. Error:" & @error & " Extended:" & @extended)

	If ($bExists = True) Then
		MsgBox($MB_OK, "", "I created a new DateTime format key.")
	Else
		MsgBox($MB_OK, "", "I Failed to create a new DateTime format key.")
	EndIf

	;Delete the newly created Format Key.
	_LOWriter_DateFormatKeyDelete($oDoc, $iFormatKey)
	If (@error > 0) Then _ERROR("Failed to delete a Date/Time Format Key. Error:" & @error & " Extended:" & @extended)

	;Check if the new key exists now.
	$bExists = _LOWriter_DateFormatKeyExists($oDoc, $iFormatKey)
	If (@error > 0) Then _ERROR("Failed to search for a Date/Time Format Key. Error:" & @error & " Extended:" & @extended)

	If ($bExists = True) Then
		MsgBox($MB_OK, "", "I failed to delete the DateTime format key.")
	Else
		MsgBox($MB_OK, "", "I successfully deleted the new DateTime format key.")
	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

