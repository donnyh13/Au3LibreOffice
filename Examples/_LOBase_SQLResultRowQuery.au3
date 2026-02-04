#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oConnection, $oTable, $oStatement, $oResult
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

	; Fill the Database with data.
	If Not _FillDatabase($oDoc, $oConnection, $oTable) Then Return

	; Create a Statement Object
	$oStatement = _LOBase_SQLStatementCreate($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to create a SQL Statement Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Execute a query, returning all columns and all entries.
	$oResult = _LOBase_SQLStatementExecuteQuery($oStatement, "SELECT * FROM ""tblNew_Table""", True)
	If @error Then Return _ERROR($oDoc, "Failed to Execute a SQL Statement Query. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Cursor to the first record.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Query if this Row has been modified
	$bReturn = _LOBase_SQLResultRowQuery($oResult, $LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED)
	If @error Then Return _ERROR($oDoc, "Failed to Query Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Has this Row been modified? True/False: " & $bReturn & @CRLF & "I will now modify the first row's value.")

	; Modify the Row
	; My first entry is for pens, modify the price from 1.99 to 2.50.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_DOUBLE, 5, 2.50)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Query if this Row has been modified
	$bReturn = _LOBase_SQLResultRowQuery($oResult, $LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED)
	If @error Then Return _ERROR($oDoc, "Failed to Query Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have modified the Row, but haven't updated it yet, does it show as being modified? True/False: " & $bReturn)

	; Update the Row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_UPDATE)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Query if this Row has been modified
	$bReturn = _LOBase_SQLResultRowQuery($oResult, $LOB_RESULT_ROW_QUERY_IS_ROW_UPDATED)
	If @error Then Return _ERROR($oDoc, "Failed to Query Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now has this Row been modified? True/False: " & $bReturn)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _FillDatabase(ByRef $oDoc, ByRef $oConnection, ByRef $oTable)
	Local $oDBase, $oColumn, $oPrepStatement
	Local $tDate

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
	_LOBase_TableColAdd($oTable, "Item", $LOB_DATA_TYPE_VARCHAR)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "In_stock", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "On_Order", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Price", $LOB_DATA_TYPE_DOUBLE)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Last_Ordered", $LOB_DATA_TYPE_DATE)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Discontinued", $LOB_DATA_TYPE_BOOLEAN)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Prepared Statement
	$oPrepStatement = _LOBase_SQLStatementCreate($oConnection, "INSERT INTO ""tblNew_Table"" (""Item"", ""In_stock"", ""On_Order"", ""Price"", ""Last_Ordered"", ""Discontinued"") VALUES (?, ?, ?, ?, ?, ?)")
	If @error Then Return _ERROR($oDoc, "Failed to create a Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 1, $LOB_DATA_SET_TYPE_STRING, "Pen")
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 2, $LOB_DATA_SET_TYPE_INT, 10)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 3, $LOB_DATA_SET_TYPE_INT, 0)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_DOUBLE, 1.99)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(2024, 5, 18)
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 5, $LOB_DATA_SET_TYPE_DATE, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 6, $LOB_DATA_SET_TYPE_BOOL, False)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Execute the Statement.
	_LOBase_SQLStatementExecuteUpdate($oPrepStatement)
	If @error Then Return _ERROR($oDoc, "Failed to Execute Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 1, $LOB_DATA_SET_TYPE_STRING, "Notebook")
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 2, $LOB_DATA_SET_TYPE_INT, 3)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 3, $LOB_DATA_SET_TYPE_INT, 17)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_DOUBLE, 12.50)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(2024, 8, 18)
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 5, $LOB_DATA_SET_TYPE_DATE, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 6, $LOB_DATA_SET_TYPE_BOOL, False)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Execute the Statement.
	_LOBase_SQLStatementExecuteUpdate($oPrepStatement)
	If @error Then Return _ERROR($oDoc, "Failed to Execute Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 1, $LOB_DATA_SET_TYPE_STRING, "Tape")
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 2, $LOB_DATA_SET_TYPE_INT, 5)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 3, $LOB_DATA_SET_TYPE_INT, 0)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_DOUBLE, 5.99)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(2023, 9, 25)
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 5, $LOB_DATA_SET_TYPE_DATE, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 6, $LOB_DATA_SET_TYPE_BOOL, False)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Execute the Statement.
	_LOBase_SQLStatementExecuteUpdate($oPrepStatement)
	If @error Then Return _ERROR($oDoc, "Failed to Execute Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 1, $LOB_DATA_SET_TYPE_STRING, "Glitter")
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 2, $LOB_DATA_SET_TYPE_INT, 9)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 3, $LOB_DATA_SET_TYPE_INT, 0)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_DOUBLE, 3.99)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(2024, 10, 7)
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 5, $LOB_DATA_SET_TYPE_DATE, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 6, $LOB_DATA_SET_TYPE_BOOL, False)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Execute the Statement.
	_LOBase_SQLStatementExecuteUpdate($oPrepStatement)
	If @error Then Return _ERROR($oDoc, "Failed to Execute Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 1, $LOB_DATA_SET_TYPE_STRING, "Balloons")
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 2, $LOB_DATA_SET_TYPE_INT, 0)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 3, $LOB_DATA_SET_TYPE_INT, 0)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_DOUBLE, 0.99)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(2022, 5, 7)
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 5, $LOB_DATA_SET_TYPE_DATE, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Statement
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 6, $LOB_DATA_SET_TYPE_BOOL, False)
	If @error Then Return _ERROR($oDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Execute the Statement.
	_LOBase_SQLStatementExecuteUpdate($oPrepStatement)
	If @error Then Return _ERROR($oDoc, "Failed to Execute Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Return True
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)

	Return False
EndFunc
