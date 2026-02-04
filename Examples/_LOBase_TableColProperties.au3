#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oTable, $oColumn
	Local $sSavePath
	Local $avSettings[0]
	Local $iFormatKey

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
	$oColumn = _LOBase_TableColAdd($oTable, "AutoIt Col", $LOB_DATA_TYPE_NUMERIC, "", "A New Number Column.")
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New Number Format Key.
	$iFormatKey = _LOBase_FormatKeyCreate($oConnection, "#,##0.000")
	If @error Then _ERROR($oDoc, "Failed to create a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Column's properties: Max character length = 55, Default value = "18", Entry is required, Decimal place in the third place,
	; don't auto insert values.
	_LOBase_TableColProperties($oConnection, $oTable, $oColumn, 55, "18", True, 3, False, $iFormatKey, $LOB_COL_TXT_ALIGN_CENTER)
	If @error Then Return _ERROR($oDoc, "Failed to set the Column's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOBase_TableColProperties($oConnection, $oTable, $oColumn)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve the Column's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Column's current settings are as follows: " & @CRLF & _
			"The Maximum number of characters that can be entered is: " & $avSettings[0] & @CRLF & _
			"The Default value is: " & $avSettings[1] & @CRLF & _
			"Is this column required to be filled in? True/False: " & $avSettings[2] & @CRLF & _
			"The Decimal place is: " & $avSettings[3] & @CRLF & _
			"Is a value automatically inserted? True/False: " & $avSettings[4] & @CRLF & _
			"The Format Key used is: " & $avSettings[5] & @CRLF & _
			"The alignment of the Column's text is (See UDF Constants): " & $avSettings[6])

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
