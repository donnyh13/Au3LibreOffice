#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oTextBox, $oTextCursor
	Local $iHMM, $iTabStop
	Local $avTabStop

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Textbox.
	$oTextBox = _LODraw_ShapeTextBoxInsert($oPage, 10500, 5000, -1, 1500)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the Textbox.
	$oTextCursor = _LODraw_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LODraw_CursorInsertString($oTextCursor, "1. Created by AutoIt!")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Paragraph alignment to Left.
	_LODraw_CursorParAlignment($oTextCursor, $LOD_PAR_ALIGN_HOR_LEFT)
	If @error Then _ERROR($oDoc, "Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(0.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Tab Stop for the demonstration.
	$iTabStop = _LODraw_CursorParTabStopCreate($oTextCursor, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the TabStop from 1/4" to 1/2" Tab Stop position, Set alignment To $LOD_PAR_TAB_ALIGN_DECIMAL,
	; and the decimal character to ASC(.) a period, ASCII value 46, Set the fill character to Asc(~) the Tilde key ASCII Value 126..
	; Since I am modifying the TabStop position, @Extended will be my new Tab Stop position.
	_LODraw_CursorParTabStopMod($oTextCursor, $iTabStop, $iHMM, $LOD_PAR_TAB_ALIGN_DECIMAL, Asc("."), Asc("~"))
	$iTabStop = @extended
	If @error Then _ERROR($oDoc, "Failed to modify Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avTabStop = _LODraw_CursorParTabStopMod($oTextCursor, $iTabStop)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Tab stop, having the position of " & $iTabStop & " has the following settings: " & @CRLF & _
			"The Current position is, in Hundredths of a Millimeter (HMM): " & $avTabStop[0] & @CRLF & _
			"The Current Alignment setting is, (See UDF constants): " & $avTabStop[1] & @CRLF & _
			"The Current Decimal Character is, in ASC value: " & $avTabStop[2] & " and looks like: " & Chr($avTabStop[2]) & @CRLF & _
			"The Current Fill Character is, in ASC value: " & $avTabStop[3] & " and looks like: " & Chr($avTabStop[3]))

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
