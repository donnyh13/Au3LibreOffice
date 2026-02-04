#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath, $sPathCopy

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

; Delete the file.
If IsString($sPathCopy) Then FileDelete($sPathCopy)

Func Example()
	Local $oDoc
	Local $sSavePath

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@DesktopDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Desktop Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created and saved a blank L.O. Base Doc to your Desktop, found at the following Path: " _
			& $sPath & @CRLF & "Press Ok to save a copy to the desktop also.")

	; Save a copy to the Desktop
	$sPathCopy = _LOBase_DocSaveCopy($oDoc, @DesktopDir & "\A Copied Doc.odb")
	If @error Then Return _ERROR($oDoc, "Failed to save a copy of the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have saved a copy of this L.O. Base Doc to your Desktop, found at the following Path: " _
			& $sPathCopy & @CRLF & "Press Ok to close the document and delete both files.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)
EndFunc
