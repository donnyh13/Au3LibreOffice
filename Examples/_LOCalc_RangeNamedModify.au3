#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell, $oNamedRange
	Local $avSettings[0], $avRowData[1]
	Local $aavData[5]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 1 ; A1
	$aavData[0] = $avRowData

	$avRowData[0] = 457 ; A2
	$aavData[1] = $avRowData

	$avRowData[0] = 18 ; A3
	$aavData[2] = $avRowData

	$avRowData[0] = 18 ; A4
	$aavData[3] = $avRowData

	$avRowData[0] = 27 ; A4
	$aavData[4] = $avRowData

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the range A1:A5 as a Named Range in the Document (Global) Scope.
	$oNamedRange = _LOCalc_RangeNamedAdd($oDoc, $oCellRange, "My_Global_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Named Ranges. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B1 Formula to "=SUM(My_Global_Named_Range)"
	_LOCalc_CellFormula($oCell, "=SUM(My_Global_Named_Range)")
	If @error Then _ERROR($oDoc, "Failed to set Cell Formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_RangeNamedModify($oDoc, $oNamedRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Named Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Named Range's current settings are as follows: " & @CRLF & _
			"The Range currently covered by this Named range is: " & $avSettings[0] & @CRLF & _
			"The Named Range name is: " & $avSettings[1] & @CRLF & _
			"The Options, if any, set for this Named Range is (as an Integer): " & $avSettings[2] & @CRLF & _
			"The Cell being Referenced by this Range is: " & _LOCalc_RangeGetAddressAsName($avSettings[3]) & @CRLF & @CRLF & _
			"I will now modify the Named Range's settings.")

	; Retrieve Cell range C3 to C7
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C3", "C7")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number Values I want in Column A.
	$avRowData[0] = 10 ; C3
	$aavData[0] = $avRowData

	$avRowData[0] = 73 ; C4
	$aavData[1] = $avRowData

	$avRowData[0] = 28 ; C5
	$aavData[2] = $avRowData

	$avRowData[0] = 32 ; C6
	$aavData[3] = $avRowData

	$avRowData[0] = 40 ; C7
	$aavData[4] = $avRowData

	; Fill the range with Data
	_LOCalc_RangeNumbers($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B1 Formula to "=SUM(My_Global_Named_Range)"
	_LOCalc_CellFormula($oCell, "=SUM(Renamed_Named_Range)")
	If @error Then _ERROR($oDoc, "Failed to set Cell Formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Cell B2.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Named Range's settings, change the Range to cover C3:C7, set the name to "Renamed_Named_Range", Set the options of Filter and Print for this range,
	; 	and set Cell B2 as the Reference Cell.
	_LOCalc_RangeNamedModify($oDoc, $oNamedRange, $oCellRange, "Renamed_Named_Range", BitOR($LOC_NAMED_RANGE_OPT_FILTER, $LOC_NAMED_RANGE_OPT_PRINT), $oCell)
	If @error Then _ERROR($oDoc, "Failed to set Named Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_RangeNamedModify($oDoc, $oNamedRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Named Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Named Range's new settings are as follows: " & @CRLF & _
			"The Range currently covered by this Named range is: " & $avSettings[0] & @CRLF & _
			"The Named Range name is: " & $avSettings[1] & @CRLF & _
			"The Options, if any, set for this Named Range is (as an Integer): " & $avSettings[2] & @CRLF & _
			"The Cell being Referenced by this Range is: " & _LOCalc_RangeGetAddressAsName($avSettings[3]) & @CRLF & @CRLF & _
			"I will now modify the Named Range's settings to be a Formula instead of a Range.")

	; Modify the Named Range's settings, change the Range from covering C3:C7, to be a formula of SUM(A1:A3) / 4, set the name to "Named_Range_Formula"
	_LOCalc_RangeNamedModify($oDoc, $oNamedRange, "SUM(A1:A3) / 4", "Named_Range_Formula")
	If @error Then _ERROR($oDoc, "Failed to set Named Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Cell B1.
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B1 Formula to "=Named_Range_Formula"
	_LOCalc_CellFormula($oCell, "=Named_Range_Formula")
	If @error Then _ERROR($oDoc, "Failed to set Cell Formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_RangeNamedModify($oDoc, $oNamedRange)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Named Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Named Range's new settings are as follows: " & @CRLF & _
			"The Range currently covered by this Named range is: " & $avSettings[0] & @CRLF & _
			"The Named Range name is: " & $avSettings[1] & @CRLF & _
			"The Options, if any, set for this Named Range is (as an Integer): " & $avSettings[2] & @CRLF & _
			"The Cell being Referenced by this Range is: " & _LOCalc_RangeGetAddressAsName($avSettings[3]))

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
