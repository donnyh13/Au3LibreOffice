#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oParStyle
	Local $avParStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I modify the Default Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate modifying a paragraph style." & @CR & "Next Line" & @CR & "Next Line" & @LF)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Retrieve the "Default Paragraph Style" object.
	$oParStyle = _LOWriter_ParStyleGetObj($oDoc, "Default Paragraph Style")
	If @error Then _ERROR("Failed to retrieve Paragraph style object. Error:" & @error & " Extended:" & @extended)

	; Set "Default Paragraph Style" Hyphenation settings to, AutoHyphen = True, Hyphen words in caps = False, Max hyphens, 20,
	; Minimum leading characters = 3, minimum trailing = 4
	_LOWriter_ParStyleHyphenation($oParStyle, True, False, 20, 3, 4)
	If @error Then _ERROR("Failed to set the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avParStyleSettings = _LOWriter_ParStyleHyphenation($oParStyle)
	If @error Then _ERROR("Failed to retrieve the Paragraph style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Paragraph's current Hyphenation settings are as follows: " & @CRLF & _
			"Automatic hyphenation? True/False: " & $avParStyleSettings[0] & @CRLF & _
			"Hyphenate words in all CAPS? True/False: " & $avParStyleSettings[1] & @CRLF & _
			"Maximum number of consecutive hyphens: " & $avParStyleSettings[2] & @CRLF & _
			"Minimum number of characters to remain before the hyphen character: " & $avParStyleSettings[3] & @CRLF & _
			"Minimum number of characters to remain after the hyphen character: " & $avParStyleSettings[4])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
