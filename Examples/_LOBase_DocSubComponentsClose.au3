#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oTable, $oQuery
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
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "Col1")
	If @error Then Return _ERROR($oDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI.
	_LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oDoc, "Failed to open Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI in Design Mode.
	_LOBase_TableUIOpenByObject($oDoc, $oConnection, $oTable, True)
	If @error Then Return _ERROR($oDoc, "Failed to open Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Query to the Document.
	$oQuery = _LOBase_QueryAddByName($oConnection, "qryAutoIt_Query", "tblNew_Table", "*")
	If @error Then Return _ERROR($oDoc, "Failed to add a Query to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Query UI.
	_LOBase_QueryUIOpenByObject($oConnection, $oQuery)
	If @error Then Return _ERROR($oDoc, "Failed to open Query UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Query UI in Design mode.
	_LOBase_QueryUIOpenByObject($oConnection, $oQuery, True)
	If @error Then Return _ERROR($oDoc, "Failed to open Query UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new form.
	_LOBase_FormCreate($oConnection, "frmAutoIt_Form", True)
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a second form.
	_LOBase_FormCreate($oConnection, "frmAutoIt_Form2", True, False)
	If @error Then Return _ERROR($oDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Report.
	_LOBase_ReportCreate($oConnection, "rptAutoIt_Report", True)
	If @error Then Return _ERROR($oDoc, "Failed to create a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have opened Tables (both in Design and viewing mode), Queries (both in Design and viewing mode), Forms (both in Design and viewing mode) and a Report in Design mode. Press ok to close them all.")

	; Close all SubComponents.
	_LOBase_DocSubComponentsClose($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to close SubComponents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the Base document.")

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
EndFunc
