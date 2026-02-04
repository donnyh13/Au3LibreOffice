#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection
	Local $sSavePath
	Local $bReturn

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

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new form.
	_LOBase_FormCreate($oConnection, "frmAutoIt_Form", False)
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Folder
	_LOBase_FormFolderCreate($oDoc, "AutoIt_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to create a form folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new form in the Folder.
	_LOBase_FormCreate($oConnection, "AutoIt_Folder/frmAutoIt_Form2", False)
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; See if a form exists with the name "frmAutoIt_Form"
	$bReturn = _LOBase_FormExists($oDoc, "frmAutoIt_Form", True)
	If @error Then Return _ERROR($oDoc, "Failed to query if a form exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a form exist in the document with the name of ""frmAutoIt_Form""? True/False. " & $bReturn)

	; See if a form exists with the name "frmAutoIt_Form2" with Exhaustive set to True.
	$bReturn = _LOBase_FormExists($oDoc, "frmAutoIt_Form2", True)
	If @error Then Return _ERROR($oDoc, "Failed to query if a form exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a form exist in the document with the name of ""frmAutoIt_Form2""? True/False. " & $bReturn)

	; See if a form exists with the name "frmAutoIt_Form" in the folder "AutoIt_Folder".
	$bReturn = _LOBase_FormExists($oDoc, "AutoIt_Folder/frmAutoIt_Form")
	If @error Then Return _ERROR($oDoc, "Failed to query if a form exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a form exist in the document with the name of ""frmAutoIt_Form"" inside the folder ""AutoIt_Folder""? True/False. " & $bReturn)

	; See if a form exists with the name "frmAutoIt_Form2" in the folder "AutoIt_Folder".
	$bReturn = _LOBase_FormExists($oDoc, "AutoIt_Folder/frmAutoIt_Form2")
	If @error Then Return _ERROR($oDoc, "Failed to query if a form exists by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does a form exist in the document with the name of ""frmAutoIt_Form2"" inside the folder ""AutoIt_Folder""? True/False. " & $bReturn)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the Base document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
