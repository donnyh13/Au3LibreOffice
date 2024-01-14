#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I set the Paragraph style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate setting the paragraph style.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now set the current paragraph style to ""Heading 1.""")

	; Set the Paragraph style to Heading 1 using the ViewCursor.
	_LOWriter_ParStyleSet($oDoc, $oViewCursor, "Heading 1")
	If @error Then _ERROR($oDoc, "Failed to set the Paragraph style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
