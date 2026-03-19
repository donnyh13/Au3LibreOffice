
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly." & @CR & "Next Line" & _
			@LF & "Next Line")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor down one line
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_DOWN, 1)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Set the paragraph at the current cursor's location Alignment settings to, Horizontal alignment = $LOW_PAR_ALIGN_HOR_JUSTIFIED,
	;Vertical alignment = $LOW_PAR_ALIGN_VERT_CENTER, Last line alignment = $LOW_PAR_LAST_LINE_JUSTIFIED,
	;Expand single word = True, Snap to grid = False, Text direction = $LOW_TXT_DIR_LR_TB
	_LOWriter_DirFrmtParAlignment($oViewCursor, $LOW_PAR_ALIGN_HOR_JUSTIFIED, $LOW_PAR_ALIGN_VERT_CENTER, $LOW_PAR_LAST_LINE_JUSTIFIED, True, False, $LOW_TXT_DIR_LR_TB)
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtParAlignment($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Paragraph Alignment settings are as follows: " & @CRLF & _
			"Horizontal alignment, (See UDF constants): " & $avSettings[0] & @CRLF & _
			"Vertical alignment, (See UDF constants): " & $avSettings[1] & @CRLF & _
			"Last line alignment, (See UDF constants): " & $avSettings[2] & @CRLF & _
			"Expand a single word on the last line? True/False: " & $avSettings[3] & @CRLF & _
			"Snap to grid, if one is active? True/False: " & $avSettings[4] & @CRLF & _
			"Text Direction, (See UDF constants): " & $avSettings[5] & @CRLF & @CRLF & _
			"Press ok to remove direct formating.")

	;Remove Direct formatting.
	_LOWriter_DirFrmtParAlignment($oViewCursor, Default, Default, Default, Default, Default, Default)
	If (@error > 0) Then _ERROR("Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



