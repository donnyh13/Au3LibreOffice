#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oTable, $oCell
	Local $iColumns, $iRows

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Table.
	$oTable = _LODraw_TableInsert($oPage, 9000, 6000, 4, 4)
	If @error Then _ERROR($oDoc, "Failed to insert a Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of Columns
	$iColumns = _LODraw_TableColumnGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve count of Table columns. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of Rows
	$iRows = _LODraw_TableRowGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve count of Table rows. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $iCol = 0 To $iColumns - 1
		For $iRow = 0 To $iRows - 1
			; Retrieve Table Cell Object
			$oCell = _LODraw_TableCellGetObjByPosition($oTable, $iCol, $iRow)
			If @error Then _ERROR($oDoc, "Failed to retrieve Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			; Set the Table Cell's text to a name.
			_LODraw_TableCellString($oCell, String("Cell " & Chr(65 + $iCol) & $iRow + 1))
			If @error Then _ERROR($oDoc, "Failed to set Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		Next
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to insert two rows at the bottom of the table.")

	_LODraw_TableRowInsert($oTable, 2, $iRow)
	If @error Then _ERROR($oDoc, "Failed to insert a Table Row. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
