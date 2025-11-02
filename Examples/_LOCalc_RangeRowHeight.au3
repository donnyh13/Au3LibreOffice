#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oRow
	Local $avHeight[0]
	Local $iMicrometers

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Row 5's Object. Remember L.O. Rows are 0 based.
	$oRow = _LOCalc_RangeRowGetObjByPosition($oSheet, 4)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Row Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2 an inch to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(0.5)
	If @error Then _ERROR($oDoc, "Failed to convert Inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Row 5's Height to 1/2 inch.
	_LOCalc_RangeRowHeight($oRow, Null, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set Row Height. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Row 5's current height settings. Return will be an array with setting values in order of Function parameters.
	$avHeight = _LOCalc_RangeRowHeight($oRow)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row Height settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Row 5's Height settings are:" & @CRLF & _
			"Is the Row's height set to optimal? True/False: " & $avHeight[0] & @CRLF & _
			"Row 5's current height is, in Micrometers: " & $avHeight[1] & @CRLF & _
			"Press Ok to set Row 5 to optimal height now.")

	; Set Row 5's Height to Optimal.
	_LOCalc_RangeRowHeight($oRow, True)
	If @error Then _ERROR($oDoc, "Failed to set Row Height. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Row 5's height settings again.
	$avHeight = _LOCalc_RangeRowHeight($oRow)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row Height settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Row 5's new Height settings are:" & @CRLF & _
			"Is the Row's height set to optimal? True/False: " & $avHeight[0] & @CRLF & _
			"Row 5's current height is, in Micrometers: " & $avHeight[1])

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
