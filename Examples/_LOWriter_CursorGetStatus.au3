
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iPage
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text." & @CR & @CR & "Some different text" & @CR & "Another Line.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Check if the ViewCursor is currently at the end of a line.
	$bReturn = _LOWriter_CursorGetStatus($oViewCursor, $LOW_CURSOR_STAT_IS_END_OF_LINE)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the ViewCursor at the end of a line? True/False: " & $bReturn & @CRLF & @CRLF & _
			"I will now move the cursor to the beginning, and test again.")

	;Move the Cursor to the beginning of the document.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move cursor. Error:" & @error & " Extended:" & @extended)

	;Check if the ViewCursor is currently at the end of a line.
	$bReturn = _LOWriter_CursorGetStatus($oViewCursor, $LOW_CURSOR_STAT_IS_END_OF_LINE)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the ViewCursor at the end of a line? True/False: " & $bReturn)

	;Retrieve the page number the Viewcursor is currently on.
	$iPage = _LOWriter_CursorGetStatus($oViewCursor, $LOW_CURSOR_STAT_GET_PAGE)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The ViewCursor is currently on page " & $iPage)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

