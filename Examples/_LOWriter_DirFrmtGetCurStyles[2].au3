
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aoParagraphs
	Local $asStyles

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "A Line of text to Test Direct Formatting with.")
	If (@error > 0) Then _ERROR("Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array of Paragraphs for the document
	$aoParagraphs = _LOWriter_ParObjCreateList($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Paragraph Objects for Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve current styles contained in the first Paragraph. Will return a 4 element Array in the following order: current paragraph style,
	;current Character style, current Page style, current Numbering Style (if one is active).
	$asStyles = _LOWriter_DirFrmtGetCurStyles($aoParagraphs[0])
	If (@error > 0) Then _ERROR("Failed to retrieve current styles in the text selection. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "These styles are currently set for the Paragraph. Some may be blank, indicating no style is set for that style " & _
			"type: " & @CRLF & _
			"Paragraph Style: " & $asStyles[0] & @CRLF & _
			"Character Style: " & $asStyles[1] & @CRLF & _
			"Page Style: " & $asStyles[2] & @CRLF & _
			"Numbering Style: " & $asStyles[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

