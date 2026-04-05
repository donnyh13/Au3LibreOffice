#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_ViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a Footnote at the end of this line. ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Footnote at the ViewCursor.
	_LOWriter_FootnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a Footnote. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Document's Footnote settings to: Number format = $LOW_NUM_STYLE_ROMAN_UPPER, Start Numbering at 5, Before the Footnote label, Place
	; a Emdash and after the footnote label place a Colon, Counting type = $LOW_FOOTNOTE_COUNT_PER_DOC
	_LOWriter_FootnoteSettingsAutoNumber($oDoc, $LOW_NUM_STYLE_ROMAN_UPPER, 5, "—", ":", $LOW_FOOTNOTE_COUNT_PER_DOC)
	If @error Then _ERROR($oDoc, "Failed to modify Footnote settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Footnote settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FootnoteSettingsAutoNumber($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Footnote settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Document's current Footnote Auto Numbering settings are as follows: " & @CRLF & _
			"The Auto Numbering Number Style used for Footnotes is, (see UDF Constants): " & $avSettings[0] & @CRLF & _
			"The number to start Footnote AutoNumbering at is: " & $avSettings[1] & @CRLF & _
			"The string before the footnote label is: " & $avSettings[2] & @CRLF & _
			"The string after the footnote label is: " & $avSettings[3] & @CRLF & _
			"The Footnote Counting type is, (see UDF Constants): " & $avSettings[4] & @CRLF & _
			"Place the Footnotes at the end of the Document? True/False: " & $avSettings[5])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
