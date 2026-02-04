#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc
	Local $sSavePath
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test if the document has a save location/Path.
	$bReturn = _LOBase_DocHasPath($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve document information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does this document have a Save location/Path? True/False: " & $bReturn)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Test again if the document has a save location/Path.
	$bReturn = _LOBase_DocHasPath($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve document information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does this document have a Save location/Path? True/False: " & $bReturn)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
