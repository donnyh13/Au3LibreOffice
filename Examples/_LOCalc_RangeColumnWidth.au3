#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oColumn
	Local $avWidth[0]
	Local $iMicrometers

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Column C's Object.
	$oColumn = _LOCalc_RangeColumnGetObjByName($oSheet, "C")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Column Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2 an inch to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(0.5)
	If @error Then _ERROR($oDoc, "Failed to convert Inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Column C's Width to 1/2 inch.
	_LOCalc_RangeColumnWidth($oColumn, Null, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set Column Width. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Column C's current Width settings. Return will be an array with setting values in order of Function parameters.
	$avWidth = _LOCalc_RangeColumnWidth($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row Width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Column C's Width settings are:" & @CRLF & _
			"Is the Column's Width set to optimal? True/False: " & $avWidth[0] & @CRLF & _
			"Column C's current Width is, in Micrometers: " & $avWidth[1] & @CRLF & _
			"Notice that the Column Width is still showing Optimal Width is True, this is the only value it will return." & @CRLF & _
			"If I set Optimal Width to True again the Column's width will return to its automatically determined value.")

	; Set Column C's Width to Optimal = True again.
	_LOCalc_RangeColumnWidth($oColumn, True)
	If @error Then _ERROR($oDoc, "Failed to set Column Width. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Column C's Width settings again.
	$avWidth = _LOCalc_RangeColumnWidth($oColumn)
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Column C's new Width settings are:" & @CRLF & _
			"Is the Column's Width set to optimal? True/False: " & $avWidth[0] & @CRLF & _
			"Column C's current Width is, in Micrometers: " & $avWidth[1])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
