
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Table, 3 rows, 5 columns
	$oTable = _LOWriter_TableCreate($oDoc, 3, 5)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Move the ViewCursor up once, which will place it into the above table.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_UP, 1, False)
	If (@error > 0) Then _ERROR("Failed to move View Cursor. Error:" & @error & " Extended:" & @extended)

	;Retrieve the bottom left Table Cell Object.
	$oCell = _LOWriter_TableGetCellByCursor($oDoc, $oTable, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended)

	;Set the Cell background color to show which cells I have retrieved the Cell Range Object for.
	_LOWriter_CellBackColor($oCell, $LOW_COLOR_BLUE, False)

	;Set Cell text String to each Cell's name.
	_LOWriter_CellString($oCell, "This is the cell the Cursor was in.")
	If (@error > 0) Then _ERROR("Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



