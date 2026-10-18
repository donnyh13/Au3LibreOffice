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

	_LOMath_MathFormatFonts($oDoc, $LOM_FMT_FONT_KIND_NUMBERS, "Arial", True, True)
	If @error Then _ERROR($oDoc, "Failed to set Formula Font settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOMath_MathFormatFonts($oDoc, $LOM_FMT_FONT_KIND_NUMBERS)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Formula settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Formula's current Font settings for Numbers are as follows: " & @CRLF & _
			"The Font to used is: " & $avSettings[0] & @CRLF & _
			"Is the Font Bold? True/False: " & $avSettings[1] & @CRLF & _
			"Is the Font italic? True/False: " & $avSettings[2])

	_LOMath_MathFormatFonts($oDoc, $LOM_FMT_FONT_KIND_FUNCTIONS, Null, True, False)
	If @error Then _ERROR($oDoc, "Failed to set Formula Font settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOMath_MathFormatFonts($oDoc, $LOM_FMT_FONT_KIND_FUNCTIONS)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Formula settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Formula's current Font settings for Functions are as follows: " & @CRLF & _
			"The Font to used is: " & $avSettings[0] & @CRLF & _
			"Is the Font Bold? True/False: " & $avSettings[1] & @CRLF & _
			"Is the Font italic? True/False: " & $avSettings[2])

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
