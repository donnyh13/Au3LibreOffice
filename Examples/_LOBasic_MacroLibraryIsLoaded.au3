#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBasic.au3"

Example()

Func Example()
	Local $oLibContainer
	Local $bIsLoaded

	$oLibContainer = _LOBasic_MacroLibContainerGetObj()
	If @error Then _ERROR("Failed to retrieve Library Container Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$bIsLoaded = _LOBasic_MacroLibraryIsLoaded($oLibContainer, "Tools")
	If @error Then _ERROR("Failed to check if a Library is loaded. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the ""Tools"" LibreOffice Library currently loaded? True/False: " & $bIsLoaded)

	_LOBasic_MacroLibraryLoad($oLibContainer, "Tools")
	If @error Then _ERROR("Failed to load a Library. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$bIsLoaded = _LOBasic_MacroLibraryIsLoaded($oLibContainer, "Tools")
	If @error Then _ERROR("Failed to check if a Library is loaded. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now is the ""Tools"" LibreOffice Library loaded? True/False: " & $bIsLoaded)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to exit the example.")

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR("Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
