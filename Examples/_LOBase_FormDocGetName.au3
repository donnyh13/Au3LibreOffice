#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oFormDoc, $oDBase, $oConnection
	Local $sSavePath, $sName, $sFullName

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new form.
	$oFormDoc = _LOBase_FormCreate($oConnection, "frmAutoIt_Form", False)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the new Form in Design Mode.
	$oFormDoc = _LOBase_FormDocOpen($oConnection, "frmAutoIt_Form", True)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to open a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Form's Name
	$sName = _LOBase_FormDocGetName($oFormDoc, False)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to retrieve Form Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Form's Full Name
	$sFullName = _LOBase_FormDocGetName($oFormDoc, True)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to retrieve Form Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have opened the form called ""frmAutoIt_Form"" in Design Mode." & @CRLF & @CRLF & _
			"The Form document's name is:" & @CRLF & $sName & @CRLF & @CRLF & _
			"The Form document's full name is:" & @CRLF & $sFullName & @CRLF & @CRLF & _
			"Press ok to close it.")

	; Close the Form Document.
	_LOBase_FormDocClose($oFormDoc, True)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to close the form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the Base document.")

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, $oFormDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $oFormDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oFormDoc) Then _LOBase_FormDocClose($oFormDoc, True)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
