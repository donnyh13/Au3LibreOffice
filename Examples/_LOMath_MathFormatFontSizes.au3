#include <MsgBoxConstants.au3>

#include "..\LibreOfficeMath.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOMath_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Math Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOMath_MathFormula($oDoc, "{ 2 over 2 } newline { 1 + 1 }")
	If @error Then _ERROR($oDoc, "Failed to set Math formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOMath_MathFormatFontSizes($oDoc, 16, 25, 67, 98, 116, 75)
	If @error Then _ERROR($oDoc, "Failed to set Formula Font Size settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOMath_MathFormatFontSizes($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Formula settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Formula's current Font Sizes are as follows: " & @CRLF & _
			"The Base Font size is, in Points: " & $avSettings[0] & @CRLF & _
			"The percentage of the Base Font size to make Text is: " & $avSettings[1] & @CRLF & _
			"The percentage of the Base Font size to make Indices is: " & $avSettings[2] & @CRLF & _
			"The percentage of the Base Font size to make Functions is: " & $avSettings[3] & @CRLF & _
			"The percentage of the Base Font size to make Operators is: " & $avSettings[4] & @CRLF & _
			"The percentage of the Base Font size to make Limits is: " & $avSettings[5])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOMath_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOMath_DocClose($oDoc, False)
	Exit
EndFunc
