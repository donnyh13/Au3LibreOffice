#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iMicrometers, $iTabStop

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly.")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Create a TabStop at 1/4" Tab Stop position, Set the fill character to Asc(~) the Tilde key ASCII Value 126.
	; Set alignment To  $LOW_TAB_ALIGN_DECIMAL, and the decimal character to ASC(.) a period, ASCII value 46.
	$iTabStop = _LOWriter_DirFrmtParTabStopCreate($oViewCursor, $iMicrometers, Asc("~"), $LOW_TAB_ALIGN_DECIMAL, Asc("."))
	If @error Then _ERROR("Failed to Create a Paragraph Tab stop. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The new Tab stop has the position of " & $iTabStop)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
