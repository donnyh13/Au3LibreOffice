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
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended)

	; Create a Folder
	_LOBase_FormFolderCreate($oDoc, "AutoIt_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to create a form folder. Error:" & @error & " Extended:" & @extended)

	; Create a new Folder in the Folder.
	_LOBase_FormFolderCreate($oDoc, "AutoIt_Folder/Folder2")
	If @error Then Return _ERROR($oDoc, "Failed to create a form folder. Error:" & @error & " Extended:" & @extended)

	; Create another new Folder in the Folder.
	_LOBase_FormFolderCreate($oDoc, "AutoIt_Folder/Folder2/Folder3")
	If @error Then Return _ERROR($oDoc, "Failed to create a form folder. Error:" & @error & " Extended:" & @extended)

	; See if a folder exists with the name "AutoIt_Folder"
	$bReturn = _LOBase_FormFolderExists($oDoc, "AutoIt_Folder", True)
	If @error Then Return _ERROR($oDoc, "Failed to query if a folder exists by name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a folder exist in the document with the name of ""AutoIt_Folder""? True/False. " & $bReturn)

	; See if a folder exists with the name "Folder3"
	$bReturn = _LOBase_FormFolderExists($oDoc, "Folder3", True)
	If @error Then Return _ERROR($oDoc, "Failed to query if a folder exists by name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a folder exist in the document with the name of ""Folder3""? True/False. " & $bReturn)

	; See if a folder exists with the name "Folder3" within "AutoIt_Folder" with Exhaustive set to False.
	$bReturn = _LOBase_FormFolderExists($oDoc, "AutoIt_Folder/Folder3", False)
	If @error Then Return _ERROR($oDoc, "Failed to query if a folder exists by name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a folder exist in the document with the name of ""Folder3"" within the folder ""AutoIt_Folder""? True/False. " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the Base document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
