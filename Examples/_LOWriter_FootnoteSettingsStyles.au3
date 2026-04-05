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
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a Footnote at the end of this line. ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Footnote at the ViewCursor.
	_LOWriter_FootnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a Footnote. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Document's Footnote Style settings to: Paragraph Style to use = "Marginalia",Skip Page style, Text area = "Definition" Character style
	; Footnote Area = "Source Text"
	_LOWriter_FootnoteSettingsStyles($oDoc, "Marginalia", Null, "Definition", "Source Text")
	If @error Then _ERROR($oDoc, "Failed to modify Footnote settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Footnote settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FootnoteSettingsStyles($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Footnote settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Document's current Footnote Style settings are as follows: " & @CRLF & _
			"The Paragraph Style to use for Footnote content is: " & $avSettings[0] & @CRLF & _
			"The Page Style to use for End of Document Footnotes is, (If there is one): " & $avSettings[1] & @CRLF & _
			"The Character Style to use for the Footnote Anchor in the Document text is: " & $avSettings[2] & @CRLF & _
			"The Character Style to use for the Footnote area is: " & $avSettings[3])

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
