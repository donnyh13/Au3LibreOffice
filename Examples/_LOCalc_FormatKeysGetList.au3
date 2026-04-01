#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oColumn
	Local $iResults
	Local $avKeys

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New Number Format Key.
	_LOCalc_FormatKeyCreate($oDoc, "#,##0.000")
	If @error Then _ERROR($oDoc, "Failed to create a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of Format Keys. With Boolean value of whether each is a User-Created key or not, search for all Format Key types.
	$avKeys = _LOCalc_FormatKeysGetList($oDoc, True, False, $LOC_FORMAT_KEYS_ALL)
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of Keys. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iResults = @extended

	; Retrieve the object for the currently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell A1's text
	_LOCalc_CellString($oCell, "Format Key Integer")
	If @error Then _ERROR($oDoc, "Failed to set Cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell B1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell B1's text
	_LOCalc_CellString($oCell, "Format Key String")
	If @error Then _ERROR($oDoc, "Failed to set Cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell C1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Cell C1's text
	_LOCalc_CellString($oCell, "Is User-Created?")
	If @error Then _ERROR($oDoc, "Failed to set Cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To $iResults - 1
		; Retrieve a Cell in Column A
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set the cell to the Format Key Integer.
		_LOCalc_CellValue($oCell, $avKeys[$i][0])
		If @error Then _ERROR($oDoc, "Failed to set Cell value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve a Cell in Column B
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 1, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set the cell to the Format Key String.
		_LOCalc_CellString($oCell, $avKeys[$i][1])
		If @error Then _ERROR($oDoc, "Failed to set Cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve a Cell in Column C
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 2, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set the cell to the Format Key String.
		_LOCalc_CellString($oCell, String($avKeys[$i][2]))
		If @error Then _ERROR($oDoc, "Failed to set Cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve Column A's Object
	$oColumn = _LOCalc_RangeColumnGetObjByName($oSheet, "A")
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Column A's width to optimal.
	_LOCalc_RangeColumnWidth($oColumn, True)
	If @error Then _ERROR($oDoc, "Failed to set Cell width to Optimal. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Column B's Object
	$oColumn = _LOCalc_RangeColumnGetObjByName($oSheet, "B")
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Column B's width to optimal.
	_LOCalc_RangeColumnWidth($oColumn, True)
	If @error Then _ERROR($oDoc, "Failed to set Cell width to Optimal. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Column C's Object
	$oColumn = _LOCalc_RangeColumnGetObjByName($oSheet, "C")
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Column C's width to optimal.
	_LOCalc_RangeColumnWidth($oColumn, True)
	If @error Then _ERROR($oDoc, "Failed to set Cell width to Optimal. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
