#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $bReturn, $bReturn2

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the range A1:A5 as a Named Range in the Document (Global) Scope.
	_LOCalc_RangeNamedAdd($oDoc, $oCellRange, "My_Global_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Named Ranges. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range C3 to E3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C3", "E3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the range C3:E3 as a Named Range for the Sheet (local) scope.
	_LOCalc_RangeNamedAdd($oSheet, $oCellRange, "A_Local_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Named Ranges. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the document contains a Named Range with the name of "A_Local_Named_Range".
	$bReturn = _LOCalc_RangeNamedExists($oDoc, "A_Local_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to query document for Named Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the document contains a Named Range with the name of "My_Global_Named_Range".
	$bReturn2 = _LOCalc_RangeNamedExists($oDoc, "My_Global_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to query document for Named Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does the document contain a Named Range with the name of ""A_Local_Named_Range""? True/False: " & $bReturn & @CRLF & @CRLF & _
			"Does the document contain a Named Range with the name of ""My_Global_Named_Range""? True/False: " & $bReturn2)

	; Check if the Sheet contains a Named Range with the name of "A_Local_Named_Range".
	$bReturn = _LOCalc_RangeNamedExists($oSheet, "A_Local_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to query document for Named Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the Sheet contains a Named Range with the name of "My_Global_Named_Range".
	$bReturn2 = _LOCalc_RangeNamedExists($oSheet, "My_Global_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to query document for Named Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does the Sheet contain a Named Range with the name of ""A_Local_Named_Range""? True/False: " & $bReturn & @CRLF & @CRLF & _
			"Does the Sheet contain a Named Range with the name of ""My_Global_Named_Range""? True/False: " & $bReturn2)

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
