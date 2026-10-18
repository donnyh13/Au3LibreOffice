#include <MsgBoxConstants.au3>

#include "..\LibreOfficeMath.au3"

Example()

Func Example()
	Local $oDoc
	Local $iAlignment

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOMath_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Math Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOMath_MathFormula($oDoc, "{ 2 over 2 } newline { 1 + 1 }")
	If @error Then _ERROR($oDoc, "Failed to set Math formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOMath_MathFormatAlignment($oDoc, $LOM_FORMULA_ALIGN_HORI_RIGHT)
	If @error Then _ERROR($oDoc, "Failed to set formula Format settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an Integer.
	$iAlignment = _LOMath_MathFormatAlignment($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current Alignment value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Formula's current Alignment value is (See UDF Constants): " & $iAlignment)

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
