
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asCharStyles

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array of Character Style names.
	$asCharStyles = _LOWriter_CharStylesGetNames($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Character style names. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now insert a list of available Character styles. There are " & @extended & " results.")

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Character Styles available in this document are:" & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asCharStyles) -1
		;Insert the character Style name.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asCharStyles[$i] & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	;Retrieve Array of Character Style names that are applied to the document
	$asCharStyles = _LOWriter_CharStylesGetNames($oDoc, False, True)
	If (@error > 0) Then _ERROR("Failed to retrieve array of Character style names. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now display a list of used Character styles. There are " & @extended & " results.")

	;Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START,1,True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Character Styles currently in use in this document are:" & @CR & @CR,True)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asCharStyles) -1
		;Insert the character Style name.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asCharStyles[$i] & @CR)
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



