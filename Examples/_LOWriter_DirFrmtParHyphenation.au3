#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the formatting settings directly.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying formatting settings directly." & @CR & "Next Line" & _
			@CR & "Next Line")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor down one line
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_DOWN, 1)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Set the paragraph at the current cursor's location Hyphenation settings to, AutoHyphen = True, Hyphen words in caps = False,
	; Max hyphens, 20, Minimum leading characters = 3, minimum trailing = 4
	_LOWriter_DirFrmtParHyphenation($oViewCursor, True, False, 20, 3, 4)
	If @error Then _ERROR("Failed to set the Selected text's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avSettings = _LOWriter_DirFrmtParHyphenation($oViewCursor)
	If @error Then _ERROR("Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Paragraph Hyphenation settings are as follows: " & @CRLF & _
			"Automatic hyphenation? True/False: " & $avSettings[0] & @CRLF & _
			"Hyphenate words in all CAPS? True/False: " & $avSettings[1] & @CRLF & _
			"Maximum number of consecutive hyphens: " & $avSettings[2] & @CRLF & _
			"Minimum number of characters to remain before the hyphen character: " & $avSettings[3] & @CRLF & _
			"Minimum number of characters to remain after the hyphen character: " & $avSettings[4] & @CRLF & @CRLF & _
			"Press ok to remove direct formating.")

	; Remove direct formatting
	_LOWriter_DirFrmtParHyphenation($oViewCursor, Null, Null, Null, Null, Null, True)
	If @error Then _ERROR("Failed to clear the selected text's direct formatting settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
