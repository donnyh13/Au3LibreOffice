#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $iRows, $iColumns

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a Table, 5 rows, 4 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 4)
	If @error Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	; Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	; Retrieve how many Rows the Table currently contains.
	$iRows = _LOWriter_TableRowGetCount($oTable)
	If @error Then _ERROR("Failed to retrieve Text Table Row count. Error:" & @error & " Extended:" & @extended)

	; Retrieve how many Columns the Table currently contains.
	$iColumns = _LOWriter_TableColumnGetCount($oTable)
	If @error Then _ERROR("Failed to retrieve Text Table Column count. Error:" & @error & " Extended:" & @extended)

	For $iRow = 0 To $iRows - 1

		For $iColumn = 0 To $iColumns - 1
			; Retrieve each cell by position in the Table.
			$oCell = _LOWriter_TableGetCellObjByPosition($oTable, $iColumn, $iRow)
			If @error Then _ERROR("Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended)

			; Set Cell text String to Cell's position.
			_LOWriter_CellString($oCell, "Column " & $iColumn & @CR & " Row " & $iRow)
			If @error Then _ERROR("Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended)
		Next
	Next

	MsgBox($MB_OK, "", "I am going to delete the 3rd column over, Column 2, in this table.")

	_LOWriter_TableColumnDelete($oTable, 2, 1)
	If @error Then _ERROR("Failed to delete Text Table Column. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I am going to delete the columns from the 2nd column over, Column 1, to the last column in this table.")

	_LOWriter_TableColumnDelete($oTable, 1, 2)
	If @error Then _ERROR("Failed to delete Text Table Column. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
