
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $iRows, $iColumns
	Local $avRows[0], $avColumns[0]

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Retrieve how many Rows the Table currently contains.
	$iRows = _LOWriter_TableRowGetCount($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table Row count. Error:" & @error & " Extended:" & @extended)

	;Retrieve how many Columns the Table currently contains.
	$iColumns = _LOWriter_TableColumnGetCount($oTable)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table Column count. Error:" & @error & " Extended:" & @extended)

	For $iRow = 0 To $iRows - 1

		For $iColumn = 0 To $iColumns - 1
			;Retrieve each cell by position in the Table.
			$oCell = _LOWriter_TableGetCellByPosition($oTable, $iColumn, $iRow)
			If (@error > 0) Then _ERROR("Failed to retrieve Text Table Cell by position. Error:" & @error & " Extended:" & @extended)

			;Set Cell text String to Cell's position.
			_LOWriter_CellString($oCell, "Column " & $iColumn & @CR & " Row " & $iRow)
			If (@error > 0) Then _ERROR("Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended)
		Next
	Next

	;There  are two ways to set the Table data, I can retrieve an array of data in the table and modify that array, or I can create all new
	;arrays to fill it with. I will demonstrate creating all new arrays this time.
	ReDim $avRows[$iRows]
	ReDim $avColumns[$iColumns]

	;This "For" will loop through the Rows array adding a Columns Array to each element once I fill the Columns array. When I add an array to
	;another array it puts a copy into the array, so I can keep re-using  the same columns array.
	For $k = 0 To $iRows - 1

		For $i = 0 To $iColumns - 1
			$avColumns[$i] = "Array Column " & $i & @CR & "Array Row " & $k
		Next
		$avRows[$k] = $avColumns
		Sleep(10)

	Next

	MsgBox($MB_OK, "", "I will now replace the existing data in the Table")

	;Now set the data
	_LOWriter_TableSetData($oTable, $avRows)
	If (@error > 0) Then _ERROR("Failed to set Text Table Data. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



