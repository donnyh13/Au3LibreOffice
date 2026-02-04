#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oCell
	Local $asCellNames
	Local $iError
	Local $sFormula

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 3 columns, 5 rows.
	$oTable = _LOWriter_TableCreate($oDoc, $oViewCursor, 3, 5)
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Cell names.
	$asCellNames = _LOWriter_TableCellsGetNames($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert Cell names
	For $i = 0 To UBound($asCellNames) - 1
		; Retrieve each cell by name as returned in the array of cell names.
		$oCell = _LOWriter_TableGetCellObjByName($oTable, $asCellNames[$i])
		If @error Then _ERROR($oDoc, "Failed to retrieve Text Table Cell by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set Cell text String to each Cell's name.
		_LOWriter_CellString($oCell, $asCellNames[$i])
		If @error Then _ERROR($oDoc, "Failed to set Text Table Cell String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve top left ("A1") Table Cell Object
	$oCell = _LOWriter_TableGetCellObjByName($oTable, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell Formula to 2 + 2
	_LOWriter_CellFormula($oCell, "2+2")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the error value
	$iError = _LOWriter_CellGetError($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table Cell formula error value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Cell Formula.
	$sFormula = _LOWriter_CellFormula($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cell's current error value is: " & $iError & @CRLF & _
			"The Cell's current formula is: " & $sFormula)

	; Set the Cell Formula to 2 + 2 with some random letters to cause an error
	_LOWriter_CellFormula($oCell, "2+2 xdv")
	If @error Then _ERROR($oDoc, "Failed to set Text Table cell formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the error value again
	$iError = _LOWriter_CellGetError($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table Cell formula error value. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Cell Formula again.
	$sFormula = _LOWriter_CellFormula($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Table cell Formula. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cell's current error value is: " & $iError & @CRLF & _
			"The Cell's current formula is: " & $sFormula)

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
