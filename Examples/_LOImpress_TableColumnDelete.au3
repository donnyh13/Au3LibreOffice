#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTable, $oCell
	Local $iColumns, $iRows

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_BLANK
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_BLANK)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Table.
	$oTable = _LOImpress_TableInsert($oSlide, 9000, 6000, 4, 4)
	If @error Then _ERROR($oDoc, "Failed to insert a Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of Columns
	$iColumns = _LOImpress_TableColumnGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve count of Table columns. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of Rows
	$iRows = _LOImpress_TableRowGetCount($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve count of Table rows. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $iCol = 0 To $iColumns - 1
		For $iRow = 0 To $iRows - 1
			; Retrieve Table Cell Object
			$oCell = _LOImpress_TableCellGetObjByPosition($oTable, $iCol, $iRow)
			If @error Then _ERROR($oDoc, "Failed to retrieve Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			; Set the Table Cell's text to a name.
			_LOImpress_TableCellString($oCell, String("Cell " & Chr(65 + $iCol) & $iRow + 1))
			If @error Then _ERROR($oDoc, "Failed to set Table cell text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
		Next
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to delete the two middle columns.")

	; Delete some Table columns.
	_LOImpress_TableColumnDelete($oTable, 1, 2)
	If @error Then _ERROR($oDoc, "Failed to delete Table Column. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
