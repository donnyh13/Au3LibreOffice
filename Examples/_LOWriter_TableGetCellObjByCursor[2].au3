#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 5 columns, 3 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 5, 3)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the ViewCursor up once, which will place it into the above table.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_UP, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move View Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the ViewCursor right once, which will place it into the second over column, I will select as I move, thus selecting a range of cells.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move View Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; When retrieving multiple cells, a cell range will be returned, a cell range is largely the same as a single cell Object,
	; but some functions don't accept a cell range.

	; Retrieve bottom left, and Bottom second over Table Cell Objects.
	$oCell = _LOWriter_TableGetCellObjByCursor($oDoc, $oTable, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell background color to show which cells I have retrieved the Cell Range Object for.
	_LOWriter_CellBackColor($oCell, $LO_COLOR_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
