#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $iRows, $iColumns
	Local $sData

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 3 columns, 5 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 3, 5)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve how many Rows the Table currently contains.
	$iRows = _LOWriter_TableRowGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Row count. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve how many Columns the Table currently contains.
	$iColumns = _LOWriter_TableColumnGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Column count. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $iRow = 0 To $iRows - 1
		For $iColumn = 0 To $iColumns - 1
			; Retrieve each cell by position in the Table.
			$oCell = _LOWriter_TableGetCellObjByPosition($oTable, $iColumn, $iRow)
			If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			; Set Cell text String to Cell's position.
			_LOWriter_CellString($oCell, "Column " & $iColumn & @CR & " Row " & $iRow)
			If @error Then _ERROR($oDoc, "Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		Next
	Next

	; Retrieve Table data, second column over (Column 1), second row down (row 1). This will return a string.
	$sData = _LOWriter_TableGetData($oTable, 1, 1)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Data. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The returned data was: " & @CRLF & $sData)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
