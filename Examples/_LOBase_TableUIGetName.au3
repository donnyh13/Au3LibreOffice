#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oTable, $oTableUI
	Local $sSavePath, $sName, $sFullName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "Col1")
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI.
	$oTableUI = _LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to open Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Table's Name
	$sName = _LOBase_TableUIGetName($oTableUI, False)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to retrieve Table Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Table's Full Name
	$sFullName = _LOBase_TableUIGetName($oTableUI, True)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to retrieve Table Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have opened the table named ""tblNew_Table"" in Viewing/ Data editing mode." & @CRLF & @CRLF & _
			"The Table document's name is:" & @CRLF & $sName & @CRLF & @CRLF & _
			"The Table document's full name is:" & @CRLF & $sFullName & @CRLF & @CRLF & _
			"Press ok to close it.")

	; Close Table UI.
	_LOBase_TableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to close Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI.
	$oTableUI = _LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable, True)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to open Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Table's Name
	$sName = _LOBase_TableUIGetName($oTableUI, False)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to retrieve Table Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Table's Full Name
	$sFullName = _LOBase_TableUIGetName($oTableUI, True)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to retrieve Table Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have opened the table named ""tblNew_Table"" in Editing mode." & @CRLF & @CRLF & _
			"The Table document's name is:" & @CRLF & $sName & @CRLF & @CRLF & _
			"The Table document's full name is:" & @CRLF & $sFullName & @CRLF & @CRLF & _
			"Press ok to close it.")

	; Close Table UI.
	_LOBase_TableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to close Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, $oTableUI, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $oTableUI, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oTableUI) Then _LOBase_TableUIClose($oTableUI)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)
EndFunc
