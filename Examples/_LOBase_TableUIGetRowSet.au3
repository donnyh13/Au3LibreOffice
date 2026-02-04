#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oTable, $oTableUI, $oRowSet
	Local $sSavePath

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
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "Col1", $LOB_DATA_TYPE_INTEGER)
	If @error Then Return _ERROR($oDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "AutoIt Col", $LOB_DATA_TYPE_VARCHAR, "")
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI.
	$oTableUI = _LOBase_TableUIOpenByName($oConnection, "tblNew_Table")
	If @error Then Return _ERROR($oDoc, "Failed to open Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have just opened the Table UI in Data entry mode, press Ok to add some Data.")

	; Retrieve the Row Set.
	$oRowSet = _LOBase_TableUIGetRowSet($oTableUI)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Table UI Row Set. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a couple rows of Data.
	For $i = 1 To 5
		; Move to a new row to insert some Data.
		_LOBase_SQLResultRowUpdate($oRowSet, $LOB_RESULT_ROW_UPDATE_MOVE_TO_INSERT)
		If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set first Column to an Integer.
		_LOBase_SQLResultRowModify($oRowSet, $LOB_RESULT_ROW_MOD_INT, 1, $i)
		If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set second Column to some Text.
		_LOBase_SQLResultRowModify($oRowSet, $LOB_RESULT_ROW_MOD_STRING, 2, "Row " & $i)
		If @error Then Return _ERROR($oDoc, "Failed to modify Result Row Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert the new row.
		_LOBase_SQLResultRowUpdate($oRowSet, $LOB_RESULT_ROW_UPDATE_INSERT)
		If @error Then Return _ERROR($oDoc, "Failed to move to Insert Result Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have finished entering Data, press ok to close the Document.")

	; Close Table UI.
	_LOBase_TableUIClose($oTableUI)
	If @error Then Return _ERROR($oDoc, "Failed to close Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
