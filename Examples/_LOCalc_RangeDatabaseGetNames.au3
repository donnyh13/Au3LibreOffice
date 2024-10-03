#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange
	Local $asDatabase[0]
	Local $sRanges

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the range A1:A5 as a Database Range.
	_LOCalc_RangeDatabaseAdd($oDoc, $oCellRange, "My AutoIt Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Database Ranges. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range C3 to E3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C3", "E3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the range C3:E3 as a Database Range.
	_LOCalc_RangeDatabaseAdd($oDoc, $oCellRange, "A New AutoIt Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Database Ranges. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A5 to C8
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A5", "C8")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the range A5:C8 as a Database Range.
	_LOCalc_RangeDatabaseAdd($oDoc, $oCellRange, "A Third AutoIt Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Database Ranges. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a list of Database ranges set for this document.
	$asDatabase = _LOCalc_RangeDatabaseGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Database Ranges list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Combine the names into a String.
	For $i = 0 To UBound($asDatabase) - 1
		$sRanges &= '"' & $asDatabase[$i] & '"' & @CRLF
	Next

	MsgBox($MB_OK, "", "The Database Range names currently set for this document are: " & @CRLF & $sRanges)

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
