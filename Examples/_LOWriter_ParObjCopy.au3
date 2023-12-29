
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oDoc2, $oViewCursor, $oPar
	Local $aoPars[0]

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "This Paragraph Contains some direct formatting that would normally be lost when copying it, unless" & _
			" I used the clipboard." & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Move the Cursor to the beginning of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, False)
	If (@error > 0) Then _ERROR("Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended)

	;Move the Cursor to the right 4 spaces and select.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 4, True)
	If (@error > 0) Then _ERROR("Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended)

	;Set the font color to $LOW_COLOR_ORANGE, and highlight to $LOW_COLOR_PURPLE
	_LOWriter_DirFrmtFontColor($oViewCursor, $LOW_COLOR_ORANGE, Null, $LOW_COLOR_PURPLE)
	If (@error > 0) Then _ERROR("Failed to set Font Color. Error:" & @error & " Extended:" & @extended)

	;Move the Cursor to the right 43 spaces.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 43, False)
	If (@error > 0) Then _ERROR("Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended)

	;Move the Cursor to the right 45 spaces.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 45, True)
	If (@error > 0) Then _ERROR("Failed to move the View Cursor. Error:" & @error & " Extended:" & @extended)

	;Underline the selected portion in red.
	_LOWriter_DirFrmtUnderLine($oViewCursor, True, $LOW_UNDERLINE_BOLD_DASH_DOT_DOT, True, $LOW_COLOR_RED)
	If (@error > 0) Then _ERROR("Failed to set underline settings. Error:" & @error & " Extended:" & @extended)

	;Create a list of Paragraph Objects
	$aoPars = _LOWriter_ParObjCreateList($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve array of paragraphs. Error:" & @error & " Extended:" & @extended)

	;Select the first Paragraph
	_LOWriter_ParObjSelect($oDoc, $aoPars[0])
	If (@error > 0) Then _ERROR("Failed to select the paragraph. Error:" & @error & " Extended:" & @extended)

	;Copy the Paragraph.
	$oPar = _LOWriter_ParObjCopy($oDoc)
	If (@error > 0) Then _ERROR("Failed to copy the paragraphs. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have now copied the paragraph, and the object pointing to this copy is stored in $oPar." & _
			" I will now open a second document and paste the paragraph in there, and it will have preserved its formatting.")

	;Create a New, visible, Blank Libre Office Document.
	$oDoc2 = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;"Paste" the paragraph object into the new document.
	_LOWriter_ParObjPaste($oDoc2, $oPar)
	If (@error > 0) Then _ERROR("Failed to paste the paragraph object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the documents.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	;Close the document.
	_LOWriter_DocClose($oDoc2, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

