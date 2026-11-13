#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBasic.au3"

Example()

Func Example()
	Local $oLibContainer
	Local $asModules
	Local $sModules = ""

	$oLibContainer = _LOBasic_MacroLibContainerGetObj()
	If @error Then _ERROR("Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$asModules = _LOBasic_MacroModulesGetNames($oLibContainer, "Tools")
	If @error Then _ERROR("Failed to retrieve array of Module names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sModules &= $asModules[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The ""Tools"" LibreOffice Library contains the following Modules:" & @CRLF & $sModules)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to exit the example.")

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR("Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
