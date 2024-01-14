#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oPageStyle, $oFooterTextCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text at the ViewCursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text for demonstration purposes. Scroll down and look in the footer." & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	; Turn Footers on for this page style.
	_LOWriter_PageStyleFooter($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn footers on for this Page Style. Error:" & @error & " Extended:" & @extended)

	; Create a text cursor for the page style footer.
	$oFooterTextCursor = _LOWriter_DocFooterGetTextCursor($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to create a TextCursor in the Page Style's Footer. Error:" & @error & " Extended:" & @extended)

	; Insert some text in the Text Cursor.
	_LOWriter_DocInsertString($oDoc, $oFooterTextCursor, "Some text in the Page's Footer." & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

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
