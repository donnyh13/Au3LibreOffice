#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oConnection, $oTable, $oTableUI, $oPrepStatement, $oStatement
	Local $sSavePath
	Local $tDate

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

	; Fill the Database with data.
	If Not _FillDatabase($oDoc, $oConnection, $oTable) Then Return

	; Create a Prepared Statement
	$oPrepStatement = _LOBase_SQLStatementCreate($oConnection, "INSERT INTO ""tblNew_Table"" (""First_Name"", ""Screen_Name"", ""Post_Count"", ""Join_Date"") VALUES (?, ?, ?, ?)")
	If @error Then Return _ERROR($oDoc, "Failed to create a Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 1, $LOB_DATA_SET_TYPE_STRING, "Jonathan")
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 2, $LOB_DATA_SET_TYPE_STRING, "Jon")
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 3, $LOB_DATA_SET_TYPE_INT, 11476)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(2003, 12, 2)
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_DATE, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Execute the Statement.
	_LOBase_SQLStatementExecuteUpdate($oPrepStatement)
	If @error Then Return _ERROR($oDoc, "Failed to Execute Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Statement Object
	$oStatement = _LOBase_SQLStatementCreate($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to create a SQL Statement Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Execute an update, adding 57 posts to the user.
	_LOBase_SQLStatementExecuteUpdate($oStatement, "UPDATE ""tblNew_Table"" SET ""Post_Count"" = ""Post_Count""+57")
	If @error Then Return _ERROR($oDoc, "Failed to Execute a SQL Statement Query. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Open the Table UI.
	$oTableUI = _LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to close and delete the document.")

	; Close the Table UI
	_LOBase_TableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, "Failed to close Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _FillDatabase(ByRef $oDoc, ByRef $oConnection, ByRef $oTable)
	Local $oDBase, $oColumn

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "ID", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Column Object.
	$oColumn = _LOBase_TableColGetObjByIndex($oTable, 0)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the column to Auto Value.
	_LOBase_TableColProperties($oConnection, $oTable, $oColumn, Null, Null, Null, Null, True)
	If @error Then Return _ERROR($oDoc, "Failed to set Column properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "First_Name", $LOB_DATA_TYPE_VARCHAR)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Screen_Name", $LOB_DATA_TYPE_VARCHAR)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Post_Count", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Join_Date", $LOB_DATA_TYPE_DATE)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Return True
EndFunc

Func _ERROR($oDoc, $sErrorText, $oTableUI = Null)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oTableUI) Then _LOBase_TableUIClose($oTableUI)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)

	Return False
EndFunc
