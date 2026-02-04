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
	Local $iCount

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

	; Create a Folder
	_LOBase_ReportFolderCreate($oDoc, "AutoIt_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to create a Report folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Folder in the Folder.
	_LOBase_ReportFolderCreate($oDoc, "AutoIt_Folder/Folder2")
	If @error Then Return _ERROR($oDoc, "Failed to create a Report folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create another new Folder in the Folder.
	_LOBase_ReportFolderCreate($oDoc, "AutoIt_Folder/Folder2/Folder3")
	If @error Then Return _ERROR($oDoc, "Failed to create a Report folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create another new Folder in AutoIt_Folder.
	_LOBase_ReportFolderCreate($oDoc, "AutoIt_Folder/Folder4")
	If @error Then Return _ERROR($oDoc, "Failed to create a Report folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of Folders contained in the Document.
	$iCount = _LOBase_ReportFoldersGetCount($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve a count of Folders. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There are " & $iCount & " Folders contained in the Document.")

	; Retrieve a count of Folders contained only in "AutoIt_Folder" Folder.
	$iCount = _LOBase_ReportFoldersGetCount($oDoc, False, "AutoIt_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to retrieve a count of Folders. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There are " & $iCount & " Folders contained in the ""AutoIt_Folder"" Folder.")

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the Base document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
