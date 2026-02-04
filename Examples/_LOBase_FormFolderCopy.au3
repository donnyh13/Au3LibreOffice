#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection
	Local $sSavePath, $sFolders = "", $sForms = ""
	Local $asFolders[0], $asForms[0]

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
	_LOBase_FormCreate($oConnection, "frmAutoIt_Form")
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Folder
	_LOBase_FormFolderCreate($oDoc, "AutoIt_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to create a form folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new form in the Folder.
	_LOBase_FormCreate($oConnection, "AutoIt_Folder/frmAutoIt_Form2", False)
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created a folder containing a form. Press ok to copy it and its contents.")

	; Copy the Folder.
	_LOBase_FormFolderCopy($oDoc, "AutoIt_Folder", "Copied_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to copy the folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Copy the Folder again.
	_LOBase_FormFolderCopy($oDoc, "AutoIt_Folder", "AutoIt_Folder/Copied_Folder2")
	If @error Then Return _ERROR($oDoc, "Failed to copy the folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Folder names.
	$asFolders = _LOBase_FormFoldersGetNames($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve array of Folder names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sFolders &= $asFolders[$i] & @CRLF
	Next

	; Retrieve an array of Form names.
	$asForms = _LOBase_FormsGetNames($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve array of form names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sForms &= $asForms[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Here is a list of Folders contained in the document." & @CRLF & $sFolders & @CRLF & _
			"Here is a list of forms contained in the document." & @CRLF & $sForms)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the Base document.")

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
