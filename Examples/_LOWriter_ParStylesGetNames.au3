
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asParStyles

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array of Paragraph Style names.
	$asParStyles = _LOWriter_ParStylesGetNames($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Paragraph style names. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now insert a list of available Paragraph styles. There are " & @extended & " results.")

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Paragraph Styles available in this document are:" & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asParStyles) -1
		;Insert the Paragraph Style name.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asParStyles[$i] & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	;Retrieve Array of Paragraph Style names that are applied to the document
	$asParStyles = _LOWriter_ParStylesGetNames($oDoc, False, True)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Paragraph style names. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now display a list of used Paragraph styles. There is " & @extended & " result.")

	;Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START,1,True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Paragraph Styles currently in use in this document are:" & @CR & @CR,True)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asParStyles) -1
		;Insert the Paragraph Style name.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asParStyles[$i] & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

