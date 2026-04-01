#include <File.au3>
#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"
#include "..\LibreOfficeWriter.au3"
#include "..\LibreOfficeCalc.au3"
#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oWriterDoc, $oCalcDoc, $oBaseDoc, $oDBase, $oConnection, $oTable, $oQuery
	Local $sSavePath
	Local $avDocs

	; Create a new Calc Document.
	$oCalcDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Writer Document.
	$oWriterDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New, visible, Blank Libre Office Document.
	$oBaseDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oBaseDoc)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oBaseDoc, $sSavePath, True)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Database Object.
	$oDBase = _LOBase_DatabaseGetObjByDoc($oBaseDoc)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Database
	$oConnection = _LOBase_DatabaseConnectionGet($oDBase)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to create a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "Col1")
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI.
	_LOBase_TableUIOpenByObject($oBaseDoc, $oConnection, $oTable)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to open Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Table UI in Design Mode.
	_LOBase_TableUIOpenByObject($oBaseDoc, $oConnection, $oTable, True)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to open Table UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Query to the Document.
	$oQuery = _LOBase_QueryAddByName($oConnection, "qryAutoIt_Query", "tblNew_Table", "*")
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to add a Query to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Query UI.
	_LOBase_QueryUIOpenByObject($oConnection, $oQuery)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to open Query UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Open the Query UI in Design mode.
	_LOBase_QueryUIOpenByObject($oConnection, $oQuery, True)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to open Query UI. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new form.
	_LOBase_FormCreate($oConnection, "frmAutoIt_Form", True)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Folder
	_LOBase_FormFolderCreate($oBaseDoc, "AutoIt_Folder")
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to create a form folder. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new form in the Folder.
	_LOBase_FormCreate($oConnection, "AutoIt_Folder/frmAutoIt_Form2", True, False)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to create a form Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Report.
	_LOBase_ReportCreate($oConnection, "rptAutoIt_Report", True)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to create a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to all opened Documents.
	$avDocs = _LO_DocConnect($LO_DOC_CONNECT_MODE_ALL)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to connect to all opened Documents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have connected to all opened Documents, I will now display an ArrayDisplay of the results." & @CRLF & _
			"The first column will contain the Document Object. The second column will contain an Integer, matching one of the $LO_DOC_TYPE_* Constants.")

	_ArrayDisplay($avDocs)

	; Close all SubComponents.
	_LOBase_DocSubComponentsClose($oBaseDoc)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to close SubComponents. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the connection.
	_LOBase_DatabaseConnectionClose($oConnection)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to close a connection to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOBase_DocClose($oBaseDoc, False)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the Calc Document.
	_LOCalc_DocClose($oCalcDoc, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Close the Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the Writer Document.
	_LOWriter_DocClose($oWriterDoc, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Close the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oBaseDoc) Then _LOBase_DocSubComponentsClose($oBaseDoc)
	If IsObj($oBaseDoc) Then _LOBase_DocClose($oBaseDoc, False)
	If IsObj($oCalcDoc) Then _LOCalc_DocClose($oCalcDoc, False)
	If IsObj($oWriterDoc) Then _LOWriter_DocClose($oWriterDoc, False)
	Exit
EndFunc
