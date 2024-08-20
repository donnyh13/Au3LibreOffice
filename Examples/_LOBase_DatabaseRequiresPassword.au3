#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDBase
	Local $bReturn

	; Retrieve the Database Object for the already included Database, "Bibliography", that comes with the LibreOffice installation.
	$oDBase = _LOBase_DatabaseGetObjByURL("Bibliography")
	If @error Then Return _ERROR("Failed to Retrieve the Database Object. Error:" & @error & " Extended:" & @extended)

	; Check if the Database requires a Password to connect to it.
	$bReturn = _LOBase_DatabaseRequiresPassword($oDBase)
	If @error Then Return _ERROR("Failed to Query Database for Password requirement. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the Database ""Bibliography"", Password protected? True/False: " & $bReturn)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
EndFunc