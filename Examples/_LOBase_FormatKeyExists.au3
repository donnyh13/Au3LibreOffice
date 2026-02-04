#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDoc, $oReportDoc, $oDBase, $oConnection
	Local $iFormatKey, $i2ndFormatKey
	Local $bExists
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

	; Create a new FormatKey H:MM:SS (Hour, Minute, Second, AM/PM, Month Day Name(Day) Year
	$iFormatKey = _LOBase_FormatKeyCreate($oReportDoc, "H:MM:SS AM/PM -- M/NNN(D)/YYYY")

	; Check if the new key exists.
	$bExists = _LOBase_FormatKeyExists($oReportDoc, $iFormatKey)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to search for a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($bExists = True) Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I created a new Date format key.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I Failed to create a new Date format key.")
	EndIf

	; Create a New Number Format Key.
	$i2ndFormatKey = _LOBase_FormatKeyCreate($oReportDoc, "#,##0.000")
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to create a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the new key exists, searching in Number Format Keys only.
	$bExists = _LOBase_FormatKeyExists($oReportDoc, $i2ndFormatKey, $LOB_FORMAT_KEYS_NUMBER)
	If @error Then _ERROR($oDoc, $oReportDoc, "Failed to search for a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($bExists = True) Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I created a new Number format key.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I Failed to create a new Number format key.")
	EndIf

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the Report Document.
	_LOBase_ReportClose($oReportDoc, True)
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to close the Report Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the Report
	_LOBase_ReportDelete($oDoc, "rptAutoIt_Report")
	If @error Then Return _ERROR($oDoc, $oReportDoc, "Failed to delete a Report. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
