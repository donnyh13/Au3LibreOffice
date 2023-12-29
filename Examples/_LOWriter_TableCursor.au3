
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell, $oTableCursor
	Local $asCellNames

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Table, 4 rows, 5 columns
	$oTable = _LOWriter_TableCreate($oDoc, 4, 5)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array of Cell names.
	$asCellNames = _LOWriter_TableGetCellNames($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table Cell names. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asCellNames) - 1
		;Retrieve each cell by name as returned in the array of cell names
		$oCell = _LOWriter_TableGetCellObjByName($oTable, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed to retrieve Text Table Cell by name. Error:" & @error & " Extended:" & @extended)

		;Set Cell text String to each Cell's name.
		_LOWriter_CellString($oCell, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended)
	Next

	;Create a Table Cursor. -- Cursor will be created in the first cell ("A1")
	$oTableCursor = _LOWriter_TableCreateCursor($oDoc, $oTable)
	If (@error > 0) Then _ERROR("Failed to create Text Table cursor. Error:" & @error & " Extended:" & @extended)

	;Move the TableCursor right Twice, which will place it in the third over cell, (Cell "C1").
	_LOWriter_CursorMove($oTableCursor, $LOW_TABLECUR_GO_RIGHT, 2)
	If (@error > 0) Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	;Move the Table Cursor to the last cell on the top row, (Cell "E1"), selecting from Cell "C1" to cell "E1"
	_LOWriter_TableCursor($oTableCursor, "E1", True)
	If (@error > 0) Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press OK To merge Cells ""C1"", ""D1"", ""E1"" together.")

	;Merge cells "C1" to "E1"
	_LOWriter_TableCursor($oTableCursor, Null, False, True)
	If (@error > 0) Then _ERROR("Failed to merge Table cells. Error:" & @error & " Extended:" & @extended)

	;Move Table Cursor to cell "B3", without selecting any cells.
	_LOWriter_TableCursor($oTableCursor, "B3", False)
	If (@error > 0) Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press OK To split cell ""B3"" into 3 divisions horizontally.")

	;Split the cell "B3" horizontally into 3 pieces
	_LOWriter_TableCursor($oTableCursor, Null, False, Null, 3, True)
	If (@error > 0) Then _ERROR("Failed to Split Table cell. Error:" & @error & " Extended:" & @extended)

	;Move Table Cursor to cell "D2", without selecting any cells.
	_LOWriter_TableCursor($oTableCursor, "D2", False)
	If (@error > 0) Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press OK To split cell ""D2"" into 2 divisions vertically.")

	;Split the cell "D2" vertically into 2 pieces
	_LOWriter_TableCursor($oTableCursor, Null, False, Null, 2, False)
	If (@error > 0) Then _ERROR("Failed to Split Table cell. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now demonstrate how cell names etc change when splitting table cells. I will attempt to split Cell ""A4"" vertically, but will fail.")

	;Move Table Cursor to cell "A4", without selecting any cells.
	_LOWriter_TableCursor($oTableCursor, "A4", False)
	If (@error > 0) Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	;Split the cell "A4" vertically into 2 pieces
	_LOWriter_TableCursor($oTableCursor, Null, False, Null, 2, False)
	If (@error > 0) Then _ERROR("Failed to Split Table cell. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Notice the wrong cell was split. That is because when you split a cell, it adds to the column/row count for each split, " & _
			"making what used to be cell ""A4"" now be called ""A7"" because we split ""B3"" into three pieces. I will try splitting that cell again, this time requesting cell ""A7"".")

	;Move Table Cursor to cell "A7", without selecting any cells.
	_LOWriter_TableCursor($oTableCursor, "A7", False)
	If (@error > 0) Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	;Split the cell "A7" vertically into 2 pieces
	_LOWriter_TableCursor($oTableCursor, Null, False, Null, 2, False)
	If (@error > 0) Then _ERROR("Failed to Split Table cell. Error:" & @error & " Extended:" & @extended)
	;demonstrate error when splitting

	MsgBox($MB_OK, "", "And if I wanted to Split cell ""E2"", I would request cell ""G2"", because we split ""D2"" into two pieces, " & _
			"and two letters after ""E"", is ""G"".")

	;Move Table Cursor to cell "G2", without selecting any cells.
	_LOWriter_TableCursor($oTableCursor, "G2", False)
	If (@error > 0) Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	;Split the cell "G2" vertically into 2 pieces
	_LOWriter_TableCursor($oTableCursor, Null, False, Null, 2, False)
	If (@error > 0) Then _ERROR("Failed to Split Table cell. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to see the new cell names.")

	;Retrieve Array of Cell names.
	$asCellNames = _LOWriter_TableGetCellNames($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table Cell names. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asCellNames) - 1
		;Retrieve each cell by name as returned in the array of cell names
		$oCell = _LOWriter_TableGetCellObjByName($oTable, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed to retrieve Text Table Cell by name. Error:" & @error & " Extended:" & @extended)

		;Set Cell text String to each Cell's name.
		_LOWriter_CellString($oCell, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

