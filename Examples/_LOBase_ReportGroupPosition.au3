#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oReportDoc, $oGroup, $oValue_Col_Group, $oLast_Col_Group, $oTable
	Local $iCount, $iCurPos1, $iCurPos2
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

	; Add a Table to the Database.
	$oTable = _LOBase_TableAdd($oConnection, "tblNew_Table", "ID")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a table to the Database. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Column to the Table.
	_LOBase_TableColAdd($oTable, "Value_Col", $LOB_DATA_TYPE_INTEGER, "", "A New Integer Column.")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add another Column to the Table.
	_LOBase_TableColAdd($oTable, "Third_Col", $LOB_DATA_TYPE_VARCHAR)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a fourth Column to the Table.
	_LOBase_TableColAdd($oTable, "Last_Col", $LOB_DATA_TYPE_BOOLEAN)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a Column to the Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Report and open it.
	$oReportDoc = _LOBase_ReportCreate($oConnection, "rptAutoIt_Report", True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to create a Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a Group
	$oValue_Col_Group = _LOBase_ReportGroupAdd($oReportDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a new Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Groups properties.
	_LOBase_ReportGroupSort($oValue_Col_Group, "Value_Col")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify the Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add another Group
	$oGroup = _LOBase_ReportGroupAdd($oReportDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a new Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Groups properties.
	_LOBase_ReportGroupSort($oGroup, "ID")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify the Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add another Group
	$oLast_Col_Group = _LOBase_ReportGroupAdd($oReportDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a new Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Groups properties.
	_LOBase_ReportGroupSort($oLast_Col_Group, "Last_Col")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify the Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add another Group
	$oGroup = _LOBase_ReportGroupAdd($oReportDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to add a new Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Groups properties.
	_LOBase_ReportGroupSort($oGroup, "Third_Col")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify the Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to move the Group ""Last_Col"" to the beginning of the list, and the Group ""Value_Col"" to the end.")

	; Move "Last_Col"" to the beginning.
	_LOBase_ReportGroupPosition($oLast_Col_Group, 0)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify the Group's position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get a count of Groups
	$iCount = _LOBase_ReportGroupsGetCount($oReportDoc)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to get a count of Groups. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move ""Value_Col""" to the end.
	_LOBase_ReportGroupPosition($oValue_Col_Group, $iCount)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify the Group's position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve their Current positions.
	$iCurPos1 = _LOBase_ReportGroupPosition($oValue_Col_Group)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve the Group's position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve their Current positions.
	$iCurPos2 = _LOBase_ReportGroupPosition($oLast_Col_Group)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve the Group's position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Column ""Last_Col"" is now at position: " & $iCurPos2 & @CRLF & _
			"The Column ""Value_Col"" is now at position: " & $iCurPos1 & @CRLF & _
			"Press Ok to close the document.")

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

Func _ERROR($oDoc, $oReportDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oReportDoc) Then _LOBase_ReportClose($oReportDoc, True)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
	Exit
EndFunc
