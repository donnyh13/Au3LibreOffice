#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDBase
	Local $bReturn

	; Retrieve the Database Object for the already included Database, "Bibliography", that comes with the LibreOffice installation.
	$oDBase = _LOBase_DatabaseGetObjByURL("Bibliography")
	If @error Then Return _ERROR("Failed to Retrieve the Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the Database is Read Only.
	$bReturn = _LOBase_DatabaseIsReadOnly($oDBase)
	If @error Then Return _ERROR("Failed to Query Database for Read-Only status. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the Database ""Bibliography"", currently set to Read Only? True/False: " & $bReturn)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
EndFunc
