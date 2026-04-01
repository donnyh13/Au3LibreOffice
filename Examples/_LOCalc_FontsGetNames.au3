#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oCell, $oSheet, $oColumn
	Local $asFonts

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array list of font names
	$asFonts = _LOCalc_FontsGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of font names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " fonts found. I will now display the results in the Calc sheet. The Array will have four columns, " & @CRLF & _
			"-the first column contains the font name, " & @CRLF & _
			"-the second column contains the style name, " & @CRLF & _
			"-the third column contains the Font weight (Bold) value, (see constants)," & @CRLF & _
			"-the fourth column contains the font slant (Italic), (See constants).")

	; Retrieve the currently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve currently active sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOCalc_CellString($oCell, "Font Name")
	If @error Then _ERROR($oDoc, "Failed to insert string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 1, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOCalc_CellString($oCell, "Font Style Name")
	If @error Then _ERROR($oDoc, "Failed to insert string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 2, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOCalc_CellString($oCell, "Font Weight")
	If @error Then _ERROR($oDoc, "Failed to insert string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 3, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_LOCalc_CellString($oCell, "Font Slant")
	If @error Then _ERROR($oDoc, "Failed to insert string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($asFonts) - 1
		; Insert the Font Name
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 0, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		_LOCalc_CellString($oCell, $asFonts[$i][0])
		If @error Then _ERROR($oDoc, "Failed to insert string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert the Style Name
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 1, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		_LOCalc_CellString($oCell, $asFonts[$i][1])
		If @error Then _ERROR($oDoc, "Failed to insert string. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert the Font Weight
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 2, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		_LOCalc_CellValue($oCell, $asFonts[$i][2])
		If @error Then _ERROR($oDoc, "Failed to insert value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert the Font slant
		$oCell = _LOCalc_RangeGetCellByPosition($oSheet, 3, $i + 1)
		If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		_LOCalc_CellValue($oCell, $asFonts[$i][3])
		If @error Then _ERROR($oDoc, "Failed to insert value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve Column A's Object
	$oColumn = _LOCalc_RangeColumnGetObjByName($oSheet, "A")
	If @error Then _ERROR($oDoc, "Failed to retrieve Column Object by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Column A's width to optimal.
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
