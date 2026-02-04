#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sSavepath

Example()

; Delete the file.
If IsString($sSavepath) Then FileDelete($sSavepath)

Func Example()
	Local $oDoc
	Local $sSaveName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "A New Base Document was successfully opened. Press OK to close and save it.")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Temporary Unique File name.
	$sSaveName = "TestCloseDocument_" & @YEAR & "_" & @MON & "_" & @YDAY & "_" & @HOUR & "_" & @MIN & "_" & @SEC

	; Close the document, save changes.
	$sSavepath = _LOBase_DocClose($oDoc, True, $sSaveName)
	If @error Then Return _ERROR($oDoc, "Failed to close and save opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Base Document was successfully saved to the following path: " & $sSavepath & @CRLF & _
			"Press OK to Delete it.")
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
