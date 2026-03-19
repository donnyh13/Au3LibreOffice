
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

	;Modify the Document's Footnote settings to: Number format = $LOW_NUM_STYLE_ROMAN_UPPER, Start Numbering at 5, Before the Footnote label, Place
	;a Emdash and after the footnote label place a Colon, Counting type = $LOW_FOOTNOTE_COUNT_PER_DOC
	_LOWriter_FootnoteSettingsAutoNumber($oDoc, $LOW_NUM_STYLE_ROMAN_UPPER, 5, "—", ":", $LOW_FOOTNOTE_COUNT_PER_DOC)
	If (@error > 0) Then _ERROR("Failed to modify Footnote settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Footnote settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FootnoteSettingsAutoNumber($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve Footnote settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Document's current Footnote Auto Numbering settings are as follows: " & @CRLF & _
			"The Auto Numbering Number Style used for Footnotes is, (see UDF Constants): " & $avSettings[0] & @CRLF & _
			"The number to start Footnote AutoNumbering at is: " & $avSettings[1] & @CRLF & _
			"The string before the footnote label is: " & $avSettings[2] & @CRLF & _
			"The string after the footnote label is: " & $avSettings[3] & @CRLF & _
			"The Footnote Counting type is, (see UDF Constants): " & $avSettings[4] & @CRLF & _
			"Place the Footnotes at the end of the Document? True/False: " & $avSettings[5])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



