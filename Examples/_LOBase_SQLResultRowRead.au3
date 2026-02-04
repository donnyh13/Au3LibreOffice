#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oConnection, $oTable, $oTableUI, $oStatement, $oResult
	Local $sSavePath
	Local $sValue

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

	; Open the Table UI.
	$oTableUI = _LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Create a Statement Object
	$oStatement = _LOBase_SQLStatementCreate($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to create a SQL Statement Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Execute a query, returning all columns and all entries.
	$oResult = _LOBase_SQLStatementExecuteQuery($oStatement, "SELECT * FROM ""tblNew_Table""", True)
	If @error Then Return _ERROR($oDoc, "Failed to Execute a SQL Statement Query. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Move the Cursor to the first record.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	For $i = 1 To 7
		; Read the current value of the column as a String.
		$sValue = _LOBase_SQLResultRowRead($oResult, $LOB_RESULT_ROW_READ_STRING, $i)
		If @error Then Return _ERROR($oDoc, "Failed to read Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "Row 1, Column Number " & $i & " contains the following value: " & $sValue)
	Next

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

Func _ERROR($oDoc, $sErrorText, $oTableUI = Null)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oTableUI) Then _LOBase_TableUIClose($oTableUI)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)

	Return False
EndFunc
