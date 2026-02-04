#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oTable, $oColumn, $oColumn2
	Local $sSavePath
	Local $aoPrimaryKey[0], $aoKeys[1]
	Local $asSettings[0]

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

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "Col1")
	If @error Then Return _ERROR($oDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	$oColumn = _LOBase_TableColAdd($oTable, "AutoIt Col", $LOB_DATA_TYPE_BOOLEAN, "", "A New Boolean Column.")
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a second Column to the Table.
	$oColumn2 = _LOBase_TableColAdd($oTable, "New Col", $LOB_DATA_TYPE_INTEGER, "", "A New Integer Column.")
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Primary key for the Table.
	$aoPrimaryKey = _LOBase_TablePrimaryKey($oTable)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve current primary key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		; Retrieve Column Properties.
		$asSettings = _LOBase_TableColDefinition($oTable, $aoPrimaryKey[$i])
		If @error Then Return _ERROR($oDoc, "Failed to retrieve Column Properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Primary Key Column name is " & $asSettings[0])
	Next

	; Set the Primary key to "AutoIt Col".
	$aoKeys[0] = $oColumn
	_LOBase_TablePrimaryKey($oTable, $aoKeys)
	If @error Then Return _ERROR($oDoc, "Failed to set Table Primary key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now modified the Table's primary key to ""AutoIt Col"" column.")

	; Retrieve the current Primary key for the Table.
	$aoPrimaryKey = _LOBase_TablePrimaryKey($oTable)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve current primary key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($aoPrimaryKey) - 1
		; Retrieve Column Properties.
		$asSettings = _LOBase_TableColDefinition($oTable, $aoPrimaryKey[$i])
		If @error Then Return _ERROR($oDoc, "Failed to retrieve Column Properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Primary Key Column name is " & $asSettings[0])
	Next

	; Set the Primary key to "AutoIt Col" and "New Col".
	ReDim $aoKeys[2]
	$aoKeys[0] = $oColumn
	$aoKeys[1] = $oColumn2
	_LOBase_TablePrimaryKey($oTable, $aoKeys)
	If @error Then Return _ERROR($oDoc, "Failed to set Table Primary key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have now modified the Table's primary keys to ""AutoIt Col"" and ""New Col"" columns.")

	; Retrieve the current Primary key for the Table.
	$aoPrimaryKey = _LOBase_TablePrimaryKey($oTable)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve current primary key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($aoPrimaryKey) - 1
		; Retrieve Column Properties.
		$asSettings = _LOBase_TableColDefinition($oTable, $aoPrimaryKey[$i])
		If @error Then Return _ERROR($oDoc, "Failed to retrieve Column Properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "One of the Primary Key Columns has the name of " & $asSettings[0])
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

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
	If IsString($sPath) Then FileDelete($sPath)
EndFunc
