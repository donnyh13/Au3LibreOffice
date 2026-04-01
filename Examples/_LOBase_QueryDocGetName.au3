#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oQuery, $oQueryDoc
	Local $sSavePath, $sName, $sFullName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Table to the Database.
	_LOBase_TableAdd($oConnection, "tblNew_Table", "Col1")
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Query to the Document.
	$oQuery = _LOBase_QueryAddByName($oConnection, "qryAutoIt_Query", "tblNew_Table", "*")
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to add a Query to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Query Document.
	$oQueryDoc = _LOBase_QueryDocOpenByObject($oConnection, $oQuery)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to open Query Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Query's Name
	$sName = _LOBase_QueryDocGetName($oQueryDoc, False)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to retrieve Query Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Query's Full Name
	$sFullName = _LOBase_QueryDocGetName($oQueryDoc, True)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to retrieve Query Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have opened the Query named ""qryAutoIt_Query"" in Viewing/ Data editing mode." & @CRLF & @CRLF & _
			"The Query document's name is:" & @CRLF & $sName & @CRLF & @CRLF & _
			"The Query document's full name is:" & @CRLF & $sFullName & @CRLF & @CRLF & _
			"Press ok to close it.")

	; Close Query Document.
	_LOBase_QueryDocClose($oQueryDoc)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to close Query Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Query Document.
	$oQueryDoc = _LOBase_QueryDocOpenByObject($oConnection, $oQuery, True)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to open Query Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Query's Name
	$sName = _LOBase_QueryDocGetName($oQueryDoc, False)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to retrieve Query Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Query's Full Name
	$sFullName = _LOBase_QueryDocGetName($oQueryDoc, True)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to retrieve Query Document's name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have opened the Query named ""qryAutoIt_Query"" in Editing mode." & @CRLF & @CRLF & _
			"The Query document's name is:" & @CRLF & $sName & @CRLF & @CRLF & _
			"The Query document's full name is:" & @CRLF & $sFullName & @CRLF & @CRLF & _
			"Press ok to close it.")

	; Close Query Document.
	_LOBase_QueryDocClose($oQueryDoc)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to close Query Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, $oQueryDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $oQueryDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oQueryDoc) Then _LOBase_QueryDocClose($oQueryDoc)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)
EndFunc
