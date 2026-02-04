#include <MsgBoxConstants.au3>
#include <File.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $bReturn

	; See if the Database "Bibliography" is registered.
	$bReturn = _LOBase_DatabaseRegisteredExists("Bibliography")
	If @error Then Return _ERROR("Failed to check for registered Database by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is there a Registered Database by the name of ""Bibliography""? True/ False: " & $bReturn)

	; See if the Database "FakeNamedDatabase" is registered.
	$bReturn = _LOBase_DatabaseRegisteredExists("FakeNamedDatabase")
	If @error Then Return _ERROR("Failed to check for registered Database by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is there a Registered Database by the name of ""FakeNamedDatabase""? True/ False: " & $bReturn)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
EndFunc
