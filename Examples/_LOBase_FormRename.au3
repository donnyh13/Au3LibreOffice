#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection
	Local $sSavePath, $sForms = ""
	Local $asForms[0]

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

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended)

	; Create a new form.
	_LOBase_FormCreate($oDoc, $oConnection, "frmAutoIt_Form", False)
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended)

	; Create a Folder
	_LOBase_FormFolderCreate($oDoc, "AutoIt_Folder")
	If @error Then Return _ERROR($oDoc, "Failed to create a form folder. Error:" & @error & " Extended:" & @extended)

	; Create a new form in the Folder.
	_LOBase_FormCreate($oDoc, $oConnection, "AutoIt_Folder/frmAutoIt_Form2", False)
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of Form names.
	$asForms = _LOBase_FormsGetNames($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve array of form names. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To @extended - 1
		$sForms &= $asForms[$i] & @CRLF
	Next

	MsgBox($MB_OK, "", "Here is a list of forms contained in the document." & @CRLF & $sForms & @CRLF & @CRLF & "I will now rename the two forms.")

	; Rename the Form "frmAutoIt_Form" to "frmFirst_Form"
	_LOBase_FormRename($oDoc, "frmAutoIt_Form", "frmFirst_Form")
	If @error Then Return _ERROR($oDoc, "Failed to rename form. Error:" & @error & " Extended:" & @extended)

	; Rename the Form "frmAutoIt_Form2" found in folder AutoIt_Folder, to "frmSecond_Form"
	_LOBase_FormRename($oDoc, "AutoIt_Folder/frmAutoIt_Form2", "frmSecond_Form")
	If @error Then Return _ERROR($oDoc, "Failed to rename form. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of Form names.
	$asForms = _LOBase_FormsGetNames($oDoc, True)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve array of form names. Error:" & @error & " Extended:" & @extended)

	$sForms = ""

	For $i = 0 To @extended - 1
		$sForms &= $asForms[$i] & @CRLF
	Next

	MsgBox($MB_OK, "", "Here is a new list of forms contained in the document." & @CRLF & $sForms)

	MsgBox($MB_OK, "", "Press ok to close the Base document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
