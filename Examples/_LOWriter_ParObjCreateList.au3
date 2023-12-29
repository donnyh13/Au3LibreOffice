#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aoParagraphs

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "First Line of text" & @CR & _
			"Second line of text." & @CR & _
			"Third line of text." & @CR & _
			"Fourth Line of Text.")
	If @error Then _ERROR("Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve Array of Paragraphs for the document
	$aoParagraphs = _LOWriter_ParObjCreateList($oViewCursor)
	If @error Then _ERROR("Failed to retrieve array of Paragraph Objects for Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "There were " & @extended & " paragraph objects returned." & _
			" As an example of what a paragraph Object can be used for, I will change the font size of the first paragraph to 22 point.")

	; An example of what I can do with a paragraph Object. Set the first paragraph's font size to 22 point.
	_LOWriter_DirFrmtFont($oDoc, $aoParagraphs[0], Null, 22)
	If @error Then _ERROR("Failed to direct format Paragraph Object. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)

	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
