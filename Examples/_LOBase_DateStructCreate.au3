#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oConnection, $oTable, $oStatement, $oResult, $oTableUI
	Local $sSavePath
	Local $tDateTime

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

	; Move the Cursor to insert a new row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure
	$tDateTime = _LOBase_DateStructCreate(Int(@YEAR), Int(@MON), Int(@MDAY))
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Row
	_LOBase_SQLResultRowModify($oResult, $LOB_DATA_SET_TYPE_DATE, 2, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Time Structure
	$tDateTime = _LOBase_DateStructCreate(Null, Null, Null, Int(@HOUR), Int(@MIN), Int(@SEC), Int(@MSEC))
	If @error Then Return _ERROR($oDoc, "Failed to create a Time Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Row
	_LOBase_SQLResultRowModify($oResult, $LOB_DATA_SET_TYPE_TIME, 3, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Timestamp Structure
	$tDateTime = _LOBase_DateStructCreate(Int(@YEAR), Int(@MON), Int(@MDAY), Int(@HOUR), Int(@MIN), Int(@SEC), Int(@MSEC))
	If @error Then Return _ERROR($oDoc, "Failed to create a Date and Time Struct. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Data into the Row
	_LOBase_SQLResultRowModify($oResult, $LOB_DATA_SET_TYPE_TIMESTAMP, 4, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the new Row
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_INSERT)
	If @error Then Return _ERROR($oDoc, "Failed to Insert new Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI.
	$oTableUI = _LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to Close and Delete the Document.")

	; Close the Table UI
	_LOBase_TableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, "Failed to close Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
	_LOBase_TableColAdd($oTable, "Date", $LOB_DATA_TYPE_DATE)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Time", $LOB_DATA_TYPE_TIME)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "TimeStamp", $LOB_DATA_TYPE_TIMESTAMP)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Return True
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)

	Return False
EndFunc
