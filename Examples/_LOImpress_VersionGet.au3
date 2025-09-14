#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $sVersionAndName, $sFullVersion, $sSimpleVersion

	; Retrieve the current full Office version number and name.
	$sVersionAndName = _LOImpress_VersionGet(False, True)
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current full Office version number.
	$sFullVersion = _LOImpress_VersionGet()
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current simple Office version number.
	$sSimpleVersion = _LOImpress_VersionGet(True)
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your current full Libre Office version, including the name is: " & $sVersionAndName & @CRLF & _
			"Your current full Libre Office version is: " & $sFullVersion & @CRLF & _
			"Your current simple Libre Office version is: " & $sSimpleVersion)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
