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
	Local $iValue
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

	; Open the Table UI.
	$oTableUI = _LOBase_DocTableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox(0, "", "Press Ok to Query the Table for all entries, and then modify some of them.")

	; Close the Table UI
	_LOBase_DocTableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, "Failed to close Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Create a Statement Object
	$oStatement = _LOBase_SQLStatementCreate($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to create a SQL Statement Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Execute a query, returning all columns and all entries.
	$oResult = _LOBase_SQLStatementExecuteQuery($oStatement, "SELECT * FROM ""tblNew_Table""", True)
	If @error Then Return _ERROR($oDoc, "Failed to Execute a SQL Statement Query. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Move the Cursor to the first record.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; My first entry is for pens, I wish to increase the price from 1.99 to 2.50.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_FLOAT, 5, 2.50)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; I change my mind, and will not increase the price until my current stock runs out. I cancel the changes.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_CANCEL_UPDATE)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Move the Cursor to the second record.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; My Second entry is for Notebooks, I received my order, so I will change On_Order to 0, and increase my stock by the amount I had ordered.
	; Read my order value.
	$iValue = _LOBase_SQLResultRowRead($oResult, $LOB_RESULT_ROW_MOD_INT, 4)
	If @error Then Return _ERROR($oDoc, "Failed to Read Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Read my current stock and add it to my order value.
	$iValue += _LOBase_SQLResultRowRead($oResult, $LOB_RESULT_ROW_MOD_INT, 3)
	If @error Then Return _ERROR($oDoc, "Failed to Read Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Set on order to 0.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_INT, 4, 0)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Insert my new stock value.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_INT, 3, $iValue)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Update the Row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_UPDATE)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Move the Cursor to the Third record.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; My Third entry is for Tape, my stock is low and I have just placed an order for 20 rolls. I will update On Order and the date accordingly.
	; Set on order to 20.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_INT, 4, 20)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Set last order to today's date.
	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(Int(@YEAR), Int(@MON), Int(@MDAY))
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the date to today's date.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_DATE, 6, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Update the Row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_UPDATE)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Move the Cursor to the Fourth record.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; My Fourth entry is for glitter, I have decided to stop carrying glitter, so I will set discontinued to True.
	; Change the Boolean to True.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_BOOL, 7, True)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Update the Row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_UPDATE)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Move the Cursor to the Fifth record.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; My fifth entry is for balloons, I have decided to stop carrying balloons, so I will Delete the row.
	; Delete the Row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_DELETE)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; I wish to insert a new product I will carry, which is Staplers, so I will move to a new row to insert my new product.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Set Item to "Staplers".
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_STRING, 2, "Stapler")
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Set In Stock to 0.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_INT, 3, 0)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Set On Order to 15.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_INT, 4, 15)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Set Price to 24.99.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_FLOAT, 5, 24.99)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Set last order to today's date.
	; Create a Date Structure
	$tDate = _LOBase_DateStructCreate(Int(@YEAR), Int(@MON), Int(@MDAY))
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the date to today's date.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_DATE, 6, $tDate)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Discontinued to False.
	_LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_BOOL, 7, False)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Insert the new row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_INSERT)
	If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber, $oTableUI)

	; Open the Table UI.
	$oTableUI = _LOBase_DocTableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox(0, "", "Here is the updated table." & @CRLF & _
			"Press Ok to Close and Delete the Document.")

	; Close the Table UI
	_LOBase_DocTableUIClose($oTableUI)
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
	_LOBase_TableColProperties($oTable, $oColumn, Null, Null, Null, Null, True)
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
	_LOBase_TableColAdd($oTable, "Price", $LOB_DATA_TYPE_FLOAT)
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
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_FLOAT, 1.99)
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
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_FLOAT, 12.50)
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
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_FLOAT, 5.99)
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
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_FLOAT, 3.99)
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
	_LOBase_SQLStatementPreparedSetData($oPrepStatement, 4, $LOB_DATA_SET_TYPE_FLOAT, 0.99)
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
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oTableUI) Then _LOBase_DocTableUIClose($oTableUI)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)

	Return False
EndFunc
