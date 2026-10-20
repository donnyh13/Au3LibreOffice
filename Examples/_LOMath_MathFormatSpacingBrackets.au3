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

	_LOMath_MathFormatSpacingBrackets($oDoc, 78, 97, True, 55)
	If @error Then _ERROR($oDoc, "Failed to set Formula Spacing settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOMath_MathFormatSpacingBrackets($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Formula settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Formula's current Spacing values are as follows: " & @CRLF & _
			"The percentage of vertical distance between the upper edge of the contents and the upper end of the brackets is: " & $avSettings[0] & @CRLF & _
			"The percentage of horizontal distance between the contents and the upper end of the brackets is: " & $avSettings[1] & @CRLF & _
			"Are all bracket types scaled? True/False: " & $avSettings[2] & @CRLF & _
			"The percentage to adjust the excess size is: " & $avSettings[3])

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
