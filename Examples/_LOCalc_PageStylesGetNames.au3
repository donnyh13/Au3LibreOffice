#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oColumn
	Local $asPageStyles

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the currently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve active sheet object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Page Style names.
	$asPageStyles = _LOCalc_PageStylesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of available Page styles. There are " & @extended & " results.")

	; Retrieve Cell A1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOCalc_CellString($oCell, "All Page Styles")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Text to Bold
	_LOCalc_CellFont($oCell, Null, Null, Null, $LOC_CHAR_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to Cell Font formatting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To (UBound($asPageStyles) - 1)
		; Retrieve the Cell's Object to insert text with.
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert the Page Style name.
		_LOCalc_CellString($oCell, $asPageStyles[$i])
		If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve Array of Page Style that are applied to the document
	$asPageStyles = _LOCalc_PageStylesGetNames($oDoc, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of used Page styles. There is " & @extended & " result.")

	; Retrieve Cell B1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOCalc_CellString($oCell, "Used Page Styles")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Text to Bold
	_LOCalc_CellFont($oCell, Null, Null, Null, $LOC_CHAR_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to Cell Font formatting. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To (UBound($asPageStyles) - 1)
		; Retrieve the Cell's Object to insert text with.
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 1, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert the Page Style name.
		_LOCalc_CellString($oCell, $asPageStyles[$i])
		If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
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
