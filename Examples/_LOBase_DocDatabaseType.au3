#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oDoc
	Local $sDBType

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "A New Base Document was successfully opened. I will now set and Retrieve the Database type.")

	; Set the Database type, leave it as Default value ("sdbc:embedded:hsqldb").
	_LOBase_DocDatabaseType($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Database type.
	$sDBType = _LOBase_DocDatabaseType($oDoc, Null)
	If @error Then _ERROR($oDoc, "Failed to Retrieve Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Base Document's Database type is: " & $sDBType)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close and save opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
