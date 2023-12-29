#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $iRows, $iColumns
	Local $avTableBreak

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create the Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
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

	; Retrieve the Table's Break settings.
	$avTableBreak = _LOWriter_TableBreak($oDoc, $oTable)
	If @error Then _ERROR("Failed to retrieve Table settings. Error:" & @error & " Extended:" & @extended)

	; Array elements will be in order of function's parameters.
	MsgBox($MB_OK, "", "The current Paragraph break settings table are: " & @CRLF & _
			"Paragraph Break Type (See UDF Constants): " & $avTableBreak[0] & @CRLF & _
			"Page Style to use after the break: " & $avTableBreak[1] & @CRLF & _
			"The page number offSet for after the break: " & $avTableBreak[2])

	; Change the Table Break settings to: Page break before the Table, $LOW_BREAK_PAGE_BEFORE,  Use the page style "Landscape" for after the break,
	; And start page numbering at 2.
	_LOWriter_TableBreak($oDoc, $oTable, $LOW_BREAK_PAGE_BEFORE, "Landscape", 2)
	If @error Then _ERROR("Failed to set TableBreak settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the third down (Row 2) settings again.
	$avTableBreak = _LOWriter_TableBreak($oDoc, $oTable)
	If @error Then _ERROR("Failed to retrieve row settings. Error:" & @error & " Extended:" & @extended)

	; Array elements will be in order of function's parameters.
	MsgBox($MB_OK, "", "The new Paragraph break settings table are: " & @CRLF & _
			"Paragraph Break Type (See UDF Constants): " & $avTableBreak[0] & @CRLF & _
			"Page Style to use after the break: " & $avTableBreak[1] & @CRLF & _
			"The page number offSet for after the break: " & $avTableBreak[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
