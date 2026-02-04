#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oDBase, $oConnection, $oReportDoc, $oSection
	Local $avControls
	Local $sControls, $sSavePath

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

	; Turn the header section on.
	_LOBase_ReportPageHeader($oReportDoc, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify Report Document Header. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Page Header Section of the Report.
	$oSection = _LOBase_ReportSectionGetObj($oReportDoc, $LOB_REP_SECTION_TYPE_PAGE_HEADER)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve Section Object of Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control in the Page Header Section.
	_LOBase_ReportConInsert($oSection, $LOB_REP_CON_TYPE_TEXT_BOX, 10500, 500, 4000, 2000)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to insert a Control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Turn the Footer section on.
	_LOBase_ReportPageFooter($oReportDoc, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify Report Document Footer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Page Footer Section of the Report.
	$oSection = _LOBase_ReportSectionGetObj($oReportDoc, $LOB_REP_SECTION_TYPE_PAGE_FOOTER)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve Section Object of Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control in the Page Footer Section.
	_LOBase_ReportConInsert($oSection, $LOB_REP_CON_TYPE_IMAGE_CONTROL, 2500, 800, 2500, 3000)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to insert a Control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Detail Section of the Report.
	$oSection = _LOBase_ReportSectionGetObj($oReportDoc, $LOB_REP_SECTION_TYPE_DETAIL)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve Section Object of Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control in the Detail Section.
	_LOBase_ReportConInsert($oSection, $LOB_REP_CON_TYPE_TEXT_BOX, 2500, 800, 2500, 3000)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to insert a Control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert another control in the Detail Section.
	_LOBase_ReportConInsert($oSection, $LOB_REP_CON_TYPE_LABEL, 5500, 800, 2500, 3000)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to insert a Control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of controls contained in the Detail section.
	$avControls = _LOBase_ReportConsGetList($oSection)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to retrieve array of Controls. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		$sControls &= (IsObj($avControls[$i][0])) ? ("[Object]") : ("[Not_An_Object]")
		$sControls &= @TAB & "Control Type (See UDF Constants): " & $avControls[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Control's contained in the Detail section are: " & @CRLF & $sControls)

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
