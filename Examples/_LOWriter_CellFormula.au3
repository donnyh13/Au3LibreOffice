
#include "..\LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $asCellNames
	Local $sFormula
	Local $nValue

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed To Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed To retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Table, 5 rows, 3 columns
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If (@error > 0) Then _ERROR("Failed To create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed To insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array of Cell names.
	$asCellNames = _LOWriter_TableGetCellNames($oTable)
	If (@error > 0) Then _ERROR("Failed To retrieve Text Table Cell names. Error:" & @error & " Extended:" & @extended)

	;Insert Cell names
	For $i = 0 To UBound($asCellNames) - 1
		;Retrieve each cell by name as returned in the array of cell names.
		$oCell = _LOWriter_TableGetCellObjByName($oTable, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed To retrieve Text Table Cell by name. Error:" & @error & " Extended:" & @extended)

		;Set Cell text String to each Cell's name.
		_LOWriter_CellString($oCell, $asCellNames[$i])
		If (@error > 0) Then _ERROR("Failed To set Text Table Cell String. Error:" & @error & " Extended:" & @extended)
	Next

	;Retrieve top left ("A1") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "A1")
	If (@error > 0) Then _ERROR("Failed To retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended)

	;Set the Cell Formula to (2 + 2) / 4 * 8
	_LOWriter_CellFormula($oCell, "(2+2)/4*8")
	If (@error > 0) Then _ERROR("Failed To set Text Table cell formula. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Cell Formula.
	$sFormula = _LOWriter_CellFormula($oCell)
	If (@error > 0) Then _ERROR("Failed To retrieve Text Table cell Formula. Error:" & @error & " Extended:" & @extended)

	;Retrieve the result of the formula.
	$nValue = _LOWriter_CellValue($oCell)
	If (@error > 0) Then _ERROR("Failed To retrieve Text Table cell Value. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Cell Formula is: " & $sFormula & " The result of which is: " & $nValue)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

