#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asPageStyles

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve Array of Page Style names.
	$asPageStyles = _LOWriter_PageStylesGetNames($oDoc)
	If @error Then _ERROR("Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now insert a list of available Page styles. There are " & @extended & " results.")

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The available Page Styles in this document are:" & @CR & @CR)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asPageStyles) -1
		; Insert the Page Style names.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asPageStyles[$i] & @CR)
		If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	; Retrieve Array of Page Style names that are applied to the document
	$asPageStyles = _LOWriter_PageStylesGetNames($oDoc, False, True)
	If @error Then _ERROR("Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now insert a list of used Page styles. There is " & @extended & " result.")

	; Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START,1,True)
	If @error Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Page Styles currently in use in this document are:" & @CR & @CR,True)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asPageStyles) -1
		; Insert the Page Style name.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asPageStyles[$i] & @CR)
		If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
