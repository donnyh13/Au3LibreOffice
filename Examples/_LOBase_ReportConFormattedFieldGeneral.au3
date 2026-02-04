#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oReportDoc, $oDBase, $oConnection, $oSection, $oControl
	Local $mFont
	Local $iFormat
	Local $avControl
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

	; Turn on the Page Header.
	_LOBase_ReportPageHeader($oReportDoc, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify Report Document Page Header. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Page Header Section of the Report.
	$oSection = _LOBase_ReportSectionGetObj($oReportDoc, $LOB_REP_SECTION_TYPE_PAGE_HEADER)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve Section Object of Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control in the Page Header Section.
	$oControl = _LOBase_ReportConInsert($oSection, $LOB_REP_CON_TYPE_TEXT_BOX, 10500, 500, 4000, 2000)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to insert a Control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOBase_FontDescCreate("Times New Roman", $LOB_WEIGHT_BOLD, $LOB_POSTURE_ITALIC, 16, $LO_COLOR_INDIGO, $LOB_UNDERLINE_BOLD, $LO_COLOR_GREEN, $LOB_STRIKEOUT_NONE, False, $LOB_RELIEF_NONE)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a Format key value for "#,###.00"
	$iFormat = _LOBase_FormatKeyCreate($oReportDoc, "#,###.00")
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to retrieve a Format key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOBase_ReportConFormattedFieldGeneral($oControl, "Renamed_AutoIt_Control", 'rpt:COUNT(1;1)=VALUE("2")', True, True, $LO_COLOR_GRAY, $mFont, $LOB_TXT_ALIGN_HORI_CENTER, $LOB_ALIGN_VERT_BOTTOM, $iFormat)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to modify the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings for the control. Return will be an Array in order of function parameters.
	$avControl = _LOBase_ReportConFormattedFieldGeneral($oControl)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to retrieve Control's property values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's current settings are: " & @CRLF & _
			"The Control's name is: " & $avControl[0] & @CRLF & _
			"The Conditional Print statement is: " & $avControl[1] & @CRLF & _
			"Are repeated values printed? True/False: " & $avControl[2] & @CRLF & _
			"Are repeated values printed on Group changes? True/False: " & $avControl[3] & @CRLF & _
			"The background color is (as a RGB Color Integer): " & $avControl[4] & @CRLF & _
			"This is the current Font settings. I'll just check if it is a Map. " & IsMap($avControl[5]) & @CRLF & _
			"The Horizontal Alignment is (See UDF Constants): " & $avControl[6] & @CRLF & _
			"The Vertical Alignment is (See UDF Constants): " & $avControl[7] & @CRLF & _
			"The format key used is: " & $avControl[8])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to close the document.")

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
