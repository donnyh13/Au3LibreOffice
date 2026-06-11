#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTable, $oCell
	Local $iColumns, $iRows, $iHMM, $iTabStop
	Local $avTabStop

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
	$oTable = _LOImpress_TableInsert($oSlide, 5000, 4000, 3, 3)
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

	; Retrieve top left Table Cell Object
	$oCell = _LOImpress_TableCellGetObjByPosition($oTable, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(0.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a TabStop at 1/4" Tab Stop position.
	$iTabStop = _LOImpress_TableCellParTabStopCreate($oCell, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the TabStop from 1/4" to 1/2" Tab Stop position, Set alignment To  $LOI_PAR_TAB_ALIGN_DECIMAL,
	; and the decimal character to ASC(.) a period, ASCII value 46, Set the fill character to Asc(~) the Tilde key ASCII Value 126..
	; Since I am modifying the TabStop position, @Extended will be my new Tab Stop position.
	_LOImpress_TableCellParTabStopMod($oCell, $iTabStop, $iHMM, $LOI_PAR_TAB_ALIGN_DECIMAL, Asc("."), Asc("~"))
	$iTabStop = @extended
	If @error Then _ERROR($oDoc, "Failed to modify Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avTabStop = _LOImpress_TableCellParTabStopMod($oCell, $iTabStop)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Tab stop, having the position of " & $iTabStop & " has the following settings: " & @CRLF & _
			"The Current position is, in Hundredths of a Millimeter (HMM): " & $avTabStop[0] & @CRLF & _
			"The Current Alignment setting is, (See UDF constants): " & $avTabStop[1] & @CRLF & _
			"The Current Decimal Character is, in ASC value: " & $avTabStop[2] & " and looks like: " & Chr($avTabStop[2]) & @CRLF & _
			"The Current Fill Character is, in ASC value: " & $avTabStop[3] & " and looks like: " & Chr($avTabStop[3]))

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
