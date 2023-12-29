#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oTableCursor, $oCell
	Local $sReturn
	Local $asCellNames

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a Table, 4 rows, 5 columns
	$oTable = _LOWriter_TableCreate($oDoc, 4, 5)
	If @error Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	; Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If @error Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	; Create a Table Cursor. -- Cursor will be created in the first cell ("A1")
	$oTableCursor = _LOWriter_TableCreateCursor($oDoc, $oTable)
	If @error Then _ERROR("Failed to create Text Table cursor. Error:" & @error & " Extended:" & @extended)

	; Retrieve Array of Cell names.
	$asCellNames = _LOWriter_TableGetCellNames($oTable)
	If @error Then _ERROR("Failed to retrieve Text Table Cell names. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asCellNames) - 1
		; Retrieve each cell by name as returned in the array of cell names
		$oCell = _LOWriter_TableGetCellObjByName($oTable, $asCellNames[$i])
		If @error Then _ERROR("Failed to retrieve Text Table Cell by name. Error:" & @error & " Extended:" & @extended)

		; Set Cell text String to each Cell's name.
		_LOWriter_CellString($oCell, $asCellNames[$i])
		If @error Then _ERROR("Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended)
	Next

	; Check what cell or cells the TableCursor is currently in.
	$sReturn = _LOWriter_CursorGetStatus($oTableCursor, $LOW_CURSOR_STAT_GET_RANGE_NAME)
	If @error Then _ERROR("Failed to retrieve the Text Cursor Status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "When the Table cursor has no cells selected, the cell the Table cursor is presently in, is returned. The Table cursor is in cell: " & _
			$sReturn)

	; Move the TableCursor right Twice, selecting cells as I go.
	_LOWriter_CursorMove($oTableCursor, $LOW_TABLECUR_GO_RIGHT, 2, True)
	If @error Then _ERROR("Failed to move Table Cursor. Error:" & @error & " Extended:" & @extended)

	; Check what cell or cells the TableCursor is currently in.
	$sReturn = _LOWriter_CursorGetStatus($oTableCursor, $LOW_CURSOR_STAT_GET_RANGE_NAME)
	If @error Then _ERROR("Failed to retrieve the Text Cursor Status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "When the Table cursor has cells selected, the beginning cell and the ending cell, are returned, separated by a colon." & @CRLF & _
			"The Table cursor has the following cell range selected: " & $sReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
