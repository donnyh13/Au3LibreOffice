#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oTable, $oColumn
	Local $sSavePath, $sNames = ""
	Local $asColumns[0]

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

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended)

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "Col1")
	If @error Then Return _ERROR($oDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "AutoIt Col", $LOB_DATA_TYPE_BOOLEAN, "", "A New Boolean Column.")
	If @error Then Return _ERROR($oDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended)

	; Retrieve a list of Columns
	$asColumns = _LOBase_TableColsGetNames($oTable)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve array of Column names. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To @extended - 1
		$sNames &= $asColumns[$i] & @CRLF
	Next

	MsgBox($MB_OK, "", "The Table contains the following columns: " & @CRLF & _
			$sNames & @CRLF & @CRLF & _
			"Press Ok to delete one column.")

	$sNames = ""

	; Retrieve the object for the first column.
	$oColumn = _LOBase_TableColGetObjByIndex($oTable, 0)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve Column Object. Error:" & @error & " Extended:" & @extended)

	; Delete the column.
	_LOBase_TableColDelete($oTable, $oColumn)
	If @error Then Return _ERROR($oDoc, "Failed to delete Column. Error:" & @error & " Extended:" & @extended)

	; Retrieve a list of Columns
	$asColumns = _LOBase_TableColsGetNames($oTable)
	If @error Then Return _ERROR($oDoc, "Failed to retrieve array of Column names. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To @extended - 1
		$sNames &= $asColumns[$i] & @CRLF
	Next

	MsgBox($MB_OK, "", "The Table contains the following columns: " & @CRLF & $sNames)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	If IsString($sPath) Then FileDelete($sPath)
EndFunc