#include <MsgBoxConstants.au3>
#include "..\AU3_LibreOffice_Inspector.au3"

_Example()

Func _Example()
	Local $oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return MsgBox($MB_OK, "Error", "Failed to create a ServiceManager Object.")

	_AU3_LO_MRI($oServiceManager)
	If @error Then Return MsgBox($MB_OK, "Error", "Failed to MRI an Object. @Error: " & @error & " @Extended: " & @extended)
EndFunc   ;==>_Example
