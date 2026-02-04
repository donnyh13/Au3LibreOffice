#include <MsgBoxConstants.au3>
#include <File.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDBase, $oDoc
	Local $sSavePath, $sNames = ""
	Local $asNames[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object using the Save path.
	$oDBase = _LOBase_DatabaseGetObjByURL($sPath)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Register the Database
	_LOBase_DatabaseRegisteredAdd($oDBase, "AutoIt_Database")
	If @error Then Return _ERROR($oDoc, "Failed to Register the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of Registered Database names.
	$asNames = _LOBase_DatabaseRegisteredGetNames()
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve array of Database Names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($asNames) - 1
		$sNames &= $asNames[$i] & @CRLF
		Sleep(10)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Registered Database Names are: " & @CRLF & $sNames)

	; Unregister the Database
	_LOBase_DatabaseRegisteredRemoveByName("AutoIt_Database")
	If @error Then Return _ERROR($oDoc, "Failed to Unregister the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
