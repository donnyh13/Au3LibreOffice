
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $sVersionAndName, $sFullVersion, $sSimpleVersion

	;Retrieve the current full Office version number and name.
	$sVersionAndName = _LOWriter_VersionGet(False, True)
	If (@error > 0) Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current full Office version number.
	$sFullVersion = _LOWriter_VersionGet()
	If (@error > 0) Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current simple Office version number.
	$sSimpleVersion = _LOWriter_VersionGet(True)
	If (@error > 0) Then _ERROR("Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Your current full Libre Office version, including the name is: " & $sVersionAndName & @CRLF & _
			"Your current full Libre Office version is: " & $sFullVersion & @CRLF & _
			"Your current simple Libre Office version is: " & $sSimpleVersion)

	MsgBox($MB_OK, "", "Press ok to close the document.")

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



