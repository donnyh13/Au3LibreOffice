
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

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Footnote at the end of this line. ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Footnote at the ViewCursor.
	_LOWriter_FootnoteInsert($oDoc, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to insert a Footnote. Error:" & @error & " Extended:" & @extended)

	;Modify the Document's Footnote Style settings to: Paragraph Style to use = "Marginalia",Skip Page style, Text area = "Definition" Character style
	;Footnote Area = "Source Text"
	_LOWriter_FootnoteSettingsStyles($oDoc, "Marginalia", Null, "Definition", "Source Text")
	If (@error > 0) Then _ERROR("Failed to modify Footnote settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Footnote settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FootnoteSettingsStyles($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve Footnote settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Document's current Footnote Style settings are as follows: " & @CRLF & _
			"The Paragraph Style to use for Footnote content is: " & $avSettings[0] & @CRLF & _
			"The Page Style to use for End of Document Footnotes is, (If there is one): " & $avSettings[1] & @CRLF & _
			"The Character Style to use for the Footnote Anchor in the Document text is: " & $avSettings[2] & @CRLF & _
			"The Character Style to use for the Footnote area is: " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



