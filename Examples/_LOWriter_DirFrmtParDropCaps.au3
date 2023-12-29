
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iMicrometers
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly." & @LF & "Next Line" & _
			@LF & "Next Line")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(0.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set the paragraph at the current cursor's location Drop cap settings to, Number of Characters to DropCap, 3, Lines to drop down, 2,
	;Spc To text To 1/4 ", whole word to False, and Character style to "Example".
	_LOWriter_DirFrmtParDropCaps($oDoc, $oViewCursor, 3, 2, $iMicrometers, False, "Example")
	If (@error > 0) Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtParDropCaps($oDoc, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Paragraph DropCap settings are as follows: " & @CRLF & _
			"How many characters are included in the DropCaps?: " & $avSettings[0] & @CRLF & _
			"How many lines will the Drop cap drop?: " & $avSettings[1] & @CRLF & _
			"How much distance between the DropCaps and the rest of the text? In micrometers: " & $avSettings[2] & @CRLF & _
			"Is the whole word DropCapped? True/False: " & $avSettings[3] & @CRLF & _
			"What character style will be used for the DropCaps, if any?: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press ok to remove direct formating.")

	;Remove direct formatting
	_LOWriter_DirFrmtParDropCaps($oDoc, $oViewCursor, Null, Null, Null, Null, Null, True)
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

