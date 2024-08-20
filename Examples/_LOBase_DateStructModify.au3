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
	Local $avReturn[0]

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

	; Fill the Database with data.
	If Not _FillDatabase($oDoc, $oConnection, $oTable) Then Return

	; Create a Statement Object
	$oStatement = _LOBase_SQLStatementCreate($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to create a SQL Statement Object. Error:" & @error & " Extended:" & @extended)

	; Execute a query, returning all columns and all entries.
	$oResult = _LOBase_SQLStatementExecuteQuery($oStatement, "SELECT * FROM ""tblNew_Table""", True)
	If @error Then Return _ERROR($oDoc, "Failed to Execute a SQL Statement Query. Error:" & @error & " Extended:" & @extended)

	; Move the Cursor to insert a new row.
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended)

	; Create a Date Structure
	$tDateTime = _LOBase_DateStructCreate(Int(@YEAR), Int(@MON), Int(@MDAY))
	If @error Then Return _ERROR($oDoc, "Failed to create a Date Struct. Error:" & @error & " Extended:" & @extended)

	; Insert Data into the Row
	_LOBase_SQLResultRowModify($oResult, $LOB_DATA_SET_TYPE_DATE, 2, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended)

	; Create a Time Structure
	$tDateTime = _LOBase_DateStructCreate(Null, Null, Null, Int(@HOUR), Int(@MIN), Int(@SEC), Int(@MSEC))
	If @error Then Return _ERROR($oDoc, "Failed to create a Time Struct. Error:" & @error & " Extended:" & @extended)

	; Insert Data into the Row
	_LOBase_SQLResultRowModify($oResult, $LOB_DATA_SET_TYPE_TIME, 3, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended)

	; Create a Timestamp Structure
	$tDateTime = _LOBase_DateStructCreate(Int(@YEAR), Int(@MON), Int(@MDAY), Int(@HOUR), Int(@MIN), Int(@SEC), Int(@MSEC))
	If @error Then Return _ERROR($oDoc, "Failed to create a Date and Time Struct. Error:" & @error & " Extended:" & @extended)

	; Insert Data into the Row
	_LOBase_SQLResultRowModify($oResult, $LOB_DATA_SET_TYPE_TIMESTAMP, 4, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to set Row data. Error:" & @error & " Extended:" & @extended)

	; Insert the new Row
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_INSERT)
	If @error Then Return _ERROR($oDoc, "Failed to Insert new Row. Error:" & @error & " Extended:" & @extended)

	; Open the Table UI.
	$oTableUI = _LOBase_DocTableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to Close Table UI and retrieve the date and time values.")

	; Close the Table UI
	_LOBase_DocTableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, "Failed to close Table User Interface. Error:" & @error & " Extended:" & @extended)

	; Create a Statement Object
	$oStatement = _LOBase_SQLStatementCreate($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to create a SQL Statement Object. Error:" & @error & " Extended:" & @extended)

	; Execute a query, returning all columns and all entries.
	$oResult = _LOBase_SQLStatementExecuteQuery($oStatement, "SELECT * FROM ""tblNew_Table""", True)
	If @error Then Return _ERROR($oDoc, "Failed to Execute a SQL Statement Query. Error:" & @error & " Extended:" & @extended)

	; Move the Cursor to the first row.
	_LOBase_SQLResultCursorMove($oResult, $LOB_RESULT_CURSOR_MOVE_NEXT)
	If @error Then Return _ERROR($oDoc, "Failed to move Result Row Cursor. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Timestamp value
	$tDateTime = _LOBase_SQLResultRowRead($oResult, $LOB_RESULT_ROW_READ_TIMESTAMP, 4)
	If @error Then Return _ERROR($oDoc, "Failed to read Result Row. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current values of the timestamp
	$avReturn = _LOBase_DateStructModify($tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Date and Time values. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Date and Time Values are: " & @CRLF & _
			"Year: " & $avReturn[0] & @CRLF & _
			"Month: " & $avReturn[1] & @CRLF & _
			"Day: " & $avReturn[2] & @CRLF & _
			"Hour: " & $avReturn[3] & @CRLF & _
			"Minute: " & $avReturn[4] & @CRLF & _
			"Second: " & $avReturn[5] & @CRLF & _
			"Nano-Seconds: " & $avReturn[6] & @CRLF & @CRLF & _
			"Press Ok to modify the year to 2023, and the month to June.")

	; Modify the timestamp
	_LOBase_DateStructModify($tDateTime, 2023, 6)
	If @error Then Return _ERROR($oDoc, "Failed to modify Date and Time values. Error:" & @error & " Extended:" & @extended)

	; Set the row to the updated values.
	$tDateTime = _LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_TIMESTAMP, 4, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will not modify the Time value in the time column.")

	; Retrieve the Timestamp value
	$tDateTime = _LOBase_SQLResultRowRead($oResult, $LOB_RESULT_ROW_READ_TIME, 3)
	If @error Then Return _ERROR($oDoc, "Failed to read Result Row. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current values of the time
	$avReturn = _LOBase_DateStructModify($tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Time values. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Time Values are: " & @CRLF & _
			"Hour: " & $avReturn[3] & @CRLF & _
			"Minute: " & $avReturn[4] & @CRLF & _
			"Second: " & $avReturn[5] & @CRLF & _
			"Nano-Seconds: " & $avReturn[6] & @CRLF & @CRLF & _
			"Press Ok to modify the Hour to 8, and the minutes to 55.")

	; Modify the timestamp
	_LOBase_DateStructModify($tDateTime, Null, Null, Null, 8, 55)
	If @error Then Return _ERROR($oDoc, "Failed to modify Date and Time values. Error:" & @error & " Extended:" & @extended)

	; Set the row to the updated values.
	$tDateTime = _LOBase_SQLResultRowModify($oResult, $LOB_RESULT_ROW_MOD_TIME, 3, $tDateTime)
	If @error Then Return _ERROR($oDoc, "Failed to modify Result Row. Error:" & @error & " Extended:" & @extended)

	; Update the row
	_LOBase_SQLResultRowUpdate($oResult, $LOB_RESULT_ROW_UPDATE_UPDATE)
	If @error Then Return _ERROR($oDoc, "Failed to Update Result Row. Error:" & @error & " Extended:" & @extended)

	; Open the Table UI.
	$oTableUI = _LOBase_DocTableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to Close and Delete the Document.")

	; Close the Table UI
	_LOBase_DocTableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, "Failed to close Table User Interface. Error:" & @error & " Extended:" & @extended)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _FillDatabase(ByRef $oDoc, ByRef $oConnection, ByRef $oTable)
	Local $oDBase, $oColumn

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended)

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "ID", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Column Object.
	$oColumn = _LOBase_TableColGetObjByIndex($oTable, 0)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended)

	; Set the column to Auto Value.
	_LOBase_TableColProperties($oTable, $oColumn, Null, Null, Null, Null, True)
	If @error Then Return _ERROR($oDoc, "Failed to set Column properties. Error:" & @error & " Extended:" & @extended)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Date", $LOB_DATA_TYPE_DATE)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Time", $LOB_DATA_TYPE_TIME)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "TimeStamp", $LOB_DATA_TYPE_TIMESTAMP)
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended)

	Return True
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)

	Return False
EndFunc
