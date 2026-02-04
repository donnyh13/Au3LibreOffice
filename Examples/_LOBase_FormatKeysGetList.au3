#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oReportDoc, $oDBase, $oConnection, $oTable, $oTableUI, $oPrepStatement
	Local $iResults
	Local $avKeys[0][2]
	Local $sSavePath

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Report and open it.
	$oReportDoc = _LOBase_ReportCreate($oConnection, "rptAutoIt_Report", True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to create a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New Number Format Key.
	_LOBase_FormatKeyCreate($oReportDoc, "#,##0.000")
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to create a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of Format Keys. With Boolean value of whether each is a User-Created key or not, search for all Format Key types.
	$avKeys = _LOBase_FormatKeysGetList($oReportDoc, True, False, $LOB_FORMAT_KEYS_ALL)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to retrieve an array of Date/Time Format Keys. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iResults = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " Format Keys found. I will now display the results in the Table. The Array will have three columns, " & @CRLF & _
			"-the first column contains the Format Key number, " & @CRLF & _
			"-the second column contains the Format Key String, " & @CRLF & _
			"-the third column contains a Boolean whether the Format Key is user-created (True) or not.")

	; Fill the Database with data.
	If Not _FillDatabase($oDoc, $oReportDoc, $oConnection, $oTable) Then Return

	; Create a Prepared Statement
	$oPrepStatement = _LOBase_SQLStatementCreate($oConnection, "INSERT INTO ""tblNew_Table"" (""FormatKey_Number"", ""FormatKey_String"", ""Is_User_Created"") VALUES (?, ?, ?)")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to create a Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To $iResults - 1
		; Insert Data into the Statement
		_LOBase_SQLStatementPreparedSetData($oPrepStatement, 1, $LOB_DATA_SET_TYPE_INT, $avKeys[$i][0])
		If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert Data into the Statement
		_LOBase_SQLStatementPreparedSetData($oPrepStatement, 2, $LOB_DATA_SET_TYPE_STRING, $avKeys[$i][1])
		If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert Data into the Statement
		_LOBase_SQLStatementPreparedSetData($oPrepStatement, 3, $LOB_DATA_SET_TYPE_BOOL, $avKeys[$i][2])
		If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to set Prepared Statement data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Execute the Statement.
		_LOBase_SQLStatementExecuteUpdate($oPrepStatement)
		If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Execute Prepared Statement. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		Sleep(IsInt(($i / 15) ? (10) : (0)))
	Next

	; Open the Table UI.
	$oTableUI = _LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to open Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to close the document.")

	; Close the Table UI
	_LOBase_TableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to close Table User Interface. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the Report Document.
	_LOBase_ReportClose($oReportDoc, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to close the Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _FillDatabase(ByRef $oDoc, ByRef $oReportDoc, ByRef $oConnection, ByRef $oTable)
	Local $oDBase, $oColumn

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "ID", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Column Object.
	$oColumn = _LOBase_TableColGetObjByIndex($oTable, 0)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the column to Auto Value.
	_LOBase_TableColProperties($oConnection, $oTable, $oColumn, Null, Null, Null, Null, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to set Column properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "FormatKey_Number", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "FormatKey_String", $LOB_DATA_TYPE_VARCHAR)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Is_User_Created", $LOB_DATA_TYPE_BOOLEAN)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Return True
EndFunc

Func _ERROR($oDoc, $oReportDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oReportDoc) Then _LOBase_ReportClose($oReportDoc, True)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
