#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iMicrometers, $iTabStop
	Local $avTabStop

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(0.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Tab Stop for the demonstration.
	$iTabStop = _LOWriter_DirFrmtParTabStopCreate($oViewCursor, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(0.5)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the TabStop from 1/4" to 1/2" Tab Stop position, Set the fill character to Asc(~) the Tilde key ASCII Value 126.
	; Set alignment To  $LOW_TAB_ALIGN_DECIMAL, and the decimal character to ASC(.) a period, ASCII value 46.
	; Since I am modifying the TabStop position, @Extended will be my new Tab Stop position.
	_LOWriter_DirFrmtParTabStopMod($oViewCursor, $iTabStop, $iMicrometers, Asc("~"), $LOW_TAB_ALIGN_DECIMAL, Asc("."))
	$iTabStop = @extended
	If @error Then _ERROR($oDoc, "Failed to modify Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avTabStop = _LOWriter_DirFrmtParTabStopMod($oViewCursor, $iTabStop)
	If @error Then _ERROR($oDoc, "Failed to retrieve Paragraph Tab stop settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Tab stop, having the position of " & $iTabStop & " has the following settings: " & @CRLF & _
			"The Current position is, in micrometers: " & $avTabStop[0] & @CRLF & _
			"The Current Fill Character is, in ASC value: " & $avTabStop[1] & " and looks like: " & Chr($avTabStop[1]) & @CRLF & _
			"The Current Alignment setting is, (See UDF constants): " & $avTabStop[2] & @CRLF & _
			"The Current Decimal Character is, in ASC value: " & $avTabStop[3] & " and looks like: " & Chr($avTabStop[3]))

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
