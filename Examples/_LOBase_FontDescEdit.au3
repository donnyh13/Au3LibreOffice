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
	Local $avControl[0], $avFont[0]
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

	; Retrieve the Detail Section of the Report.
	$oSection = _LOBase_ReportSectionGetObj($oReportDoc, $LOB_REP_SECTION_TYPE_DETAIL)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to retrieve Section Object of Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a control in the Detail Section.
	$oControl = _LOBase_ReportConInsert($oSection, $LOB_REP_CON_TYPE_LABEL, 10500, 500, 4500, 1500)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to insert a Control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set some Label Control settings.
	_LOBase_ReportConLabelGeneral($oControl, "Label_Control", "This is an AutoIt_Label")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to modify Control properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Font Descriptor.
	$mFont = _LOBase_FontDescCreate("Times New Roman", $LOB_WEIGHT_BOLD, $LOB_POSTURE_ITALIC, 18, $LO_COLOR_BRICK, $LOB_UNDERLINE_BOLD, $LO_COLOR_GREEN, $LOB_STRIKEOUT_NONE, True, $LOB_RELIEF_NONE, $LOB_CASEMAP_TITLE, False, True, True)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to create a Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Control's General properties.
	_LOBase_ReportConLabelGeneral($oControl, Null, Null, Null, Null, Null, Null, $mFont)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to modify a Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the properties of the control to modify the font.
	$avControl = _LOBase_ReportConLabelGeneral($oControl)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to retrieve the Control's properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Font's current settings. Return will be an Array in order of function parameters.
	$avFont = _LOBase_FontDescEdit($avControl[6])
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to retrieve the Font Descriptor's current values. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Font Descriptor's current settings are: " & @CRLF & _
			"The Name of the font used is: " & $avFont[0] & @CRLF & _
			"The Font weight is (See UDF Constants): " & $avFont[1] & @CRLF & _
			"The Font Italic setting is (See UDF Constants): " & $avFont[2] & @CRLF & _
			"The Font size is: " & $avFont[3] & @CRLF & _
			"The Font color is (as a RGB Color Integer): " & $avFont[4] & @CRLF & _
			"The Font underline style is (See UDF Constants): " & $avFont[5] & @CRLF & _
			"The Font Underline color is  (as a RGB Color Integer): " & $avFont[6] & @CRLF & _
			"The Strikeout line style is (See UDF Constants): " & $avFont[7] & @CRLF & _
			"Are individual words underlined? True/False: " & $avFont[8] & @CRLF & _
			"The Relief style is: (See UDF Constants) " & $avFont[9] & @CRLF & _
			"The Case style is: (See UDF Constants) " & $avFont[10] & @CRLF & _
			"Are the characters Hidden? True/False: " & $avFont[11] & @CRLF & _
			"Are the characters Outlined? True/False: " & $avFont[12] & @CRLF & _
			"Are the characters Shadowed? True/False: " & $avFont[13] & @CRLF & @CRLF & _
			"Press ok to modify the Font for this Label control.")

	; Modify the Font Descriptor.
	_LOBase_FontDescEdit($avControl[6], "Arial", $LOB_WEIGHT_NORMAL, $LOB_POSTURE_NONE, 16, $LO_COLOR_LIME, $LOB_UNDERLINE_DBL_WAVE, $LO_COLOR_PURPLE, Null, False, $LOB_RELIEF_ENGRAVED)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to modify the Font Descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new Font descriptor to the Label.
	_LOBase_ReportConLabelGeneral($oControl, Null, Null, Null, Null, Null, Null, $avControl[6])
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to modify the Label control. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the Report Document.
	_LOBase_ReportClose($oReportDoc, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to close the Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
