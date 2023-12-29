#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $sVersionAndName, $sFullVersion, $sSimpleVersion

	; Retrieve the current full Office version number and name.
	$sVersionAndName = _LOCalc_VersionGet(False, True)
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current full Office version number.
	$sFullVersion = _LOCalc_VersionGet()
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current simple Office version number.
	$sSimpleVersion = _LOCalc_VersionGet(True)
	If @error Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Your current full Libre Office version, including the name is: " & $sVersionAndName & @CRLF & _
			"Your current full Libre Office version is: " & $sFullVersion & @CRLF & _
			"Your current simple Libre Office version is: " & $sSimpleVersion)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
