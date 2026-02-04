#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDBase, $oConnection

	; Retrieve the Database Object for the already included Database, "Bibliography", that comes with the LibreOffice installation.
	$oDBase = _LOBase_DatabaseGetObjByURL("Bibliography")
	If @error Then Return _ERROR($oConnection, "Failed to Retrieve the Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oConnection, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have connected to the Database ""Bibliography"", it contains " & _LOBase_TablesGetCount($oConnection) & " Table.")

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oConnection, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oConnection, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oConnection) Then _LOBase_DatabaseConnectionClose($oConnection)
EndFunc
