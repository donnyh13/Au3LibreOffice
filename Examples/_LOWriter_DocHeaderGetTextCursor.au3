
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oPageStyle, $oHeaderTextCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text for demonstration purposes." & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	;Turn Headers on for this page style.
	_LOWriter_PageStyleHeader($oPageStyle, True)
	If (@error > 0) Then _ERROR("Failed to turn Headers on for this Page Style. Error:" & @error & " Extended:" & @extended)

	;Create a text cursor for the page style headerer.
	$oHeaderTextCursor = _LOWriter_DocHeaderGetTextCursor($oPageStyle, True)
	If (@error > 0) Then _ERROR("Failed to create a TextCursor in the Page Style's Header. Error:" & @error & " Extended:" & @extended)

	;Insert some text in the Text Cursor.
	_LOWriter_DocInsertString($oDoc, $oHeaderTextCursor, "Some text in the Page's Header." & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

