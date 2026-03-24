#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oRow
	Local $abSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Row 3's Object. Remember L.O. Rows are 0 based.
	$oRow = _LOCalc_RangeRowGetObjByPosition($oSheet, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Row Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Row page break
	_LOCalc_RangeRowPageBreak($oRow, True)
	If @error Then _ERROR($oDoc, "Failed to set Row Page Break Settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have inserted a manual Page Break at Row 3.")

	; Retrieve the Page Break Settings for Row 3. Return will be an array with setting values in order of Function parameters.
	$abSettings = _LOCalc_RangeRowPageBreak($oRow)
	If @error Then _ERROR($oDoc, "Failed to retrieve Row Page Break Settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Row 3's current Page Break settings are:" & @CRLF & _
			"Is there a Manual Page Break at Row 3? True/False: " & $abSettings[0] & @CRLF & _
			"Is Row 3 the start of a new Page? True/False: " & $abSettings[1])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

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
