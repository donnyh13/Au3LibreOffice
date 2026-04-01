#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $iFormatKey
	Local $bExists

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new DateFormatKey H:MM:SS (Hour, Minute, Second, AM/PM, Month Day Name(Day) Year
	$iFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "H:MM:SS AM/PM -- M/NNN(D)/YYYY")
	If @error Then _ERROR($oDoc, "Failed to create a Date/Time Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I created a new DateTime format key. Press Ok to delete it now.")

	; Delete the newly created Format Key.
	_LOWriter_DateFormatKeyDelete($oDoc, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to delete a Date/Time Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the new key exists.
	$bExists = _LOWriter_DateFormatKeyExists($oDoc, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to search for a Date/Time Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($bExists = True) Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I failed to delete the DateTime format key.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I successfully deleted the new DateTime format key.")
	EndIf

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

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
