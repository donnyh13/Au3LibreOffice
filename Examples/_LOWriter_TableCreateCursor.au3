
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oTableCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I am going to insert a table so that I can demonstrate a Table Cursor.")

	;Create a Table, 5 rows, 4 columns
	$oTable = _LOWriter_TableCreate($oDoc, 5, 4)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Create a Table Cursor at cell A1, which is the first cell of the table.
	$oTableCursor = _LOWriter_TableCreateCursor($oDoc, $oTable, "A1")
	If (@error > 0) Then _ERROR("Failed to create Text Table Cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I am going to move the table cursor down three cells, selecting them, and then combine them. In terms of usefulness, this " & _
			"is about all it is good for, is for combining or splitting table cells.")

	;Move the Table Cursor down three cells, selecting them
	_LOWriter_CursorMove($oTableCursor, $LOW_TABLECUR_GO_DOWN, 3, True)
	If (@error > 0) Then _ERROR("Failed to move Text Table Cursor. Error:" & @error & " Extended:" & @extended)

	;Merge the Cells selected by this Table Cursor.
	_LOWriter_TableCursor($oTableCursor, Null, False, True)
	If (@error > 0) Then _ERROR("Failed to merge cells selected by Text Table Cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

