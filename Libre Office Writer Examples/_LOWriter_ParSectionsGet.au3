
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>
#include <Array.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aoParagraphs, $aoSections

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "First Line of text" & @CR & _
			"Second line of text." & @CR & _
			"Third line of text." & @CR & _
			"Fourth Line of Text.")
	If (@error > 0) Then _ERROR("Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Move the Cursor to start of document.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Failed to move View Cursor. Error:" & @error & " Extended:" & @extended)

	;Move cursor right 10 spaces.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 10)
	If (@error > 0) Then _ERROR("Failed to move View Cursor. Error:" & @error & " Extended:" & @extended)

	;Insert a footnote to demonstrate a different section type.
	_LOWriter_FootnoteInsert($oDoc, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to insert a footnote. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array of Paragraphs for the document
	$aoParagraphs = _LOWriter_ParGetObjects($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Paragraph Objects for Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve Paragraph sections for the first paragraph.
	$aoSections = _LOWriter_ParSectionsGet($aoParagraphs[0])
	If (@error > 0) Then _ERROR("Failed to retrieve array of Paragraph Object sections. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "There were " & @extended & " paragraph sections returned." & _
			" As an example of what a paragraph section can be used for, I will change the font size of the first paragraph section to 22 point.")

	;An example of what I can do with a paragraph section Object. Set the first paragraph section's font size to 22 point. The Object
	;is stored in the first column [0] of the array.
	_LOWriter_DirFrmtFont($oDoc, $aoSections[0][0], Null, 22)
	If (@error > 0) Then _ERROR("Failed to direct format Paragraph Object. Error:" & @error & " Extended:" & @extended)

	;Display the paragraph sections.
	_ArrayDisplay($aoSections)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)

	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



