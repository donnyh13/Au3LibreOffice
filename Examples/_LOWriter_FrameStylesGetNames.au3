
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $asNames

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new FrameStyle named "Test Style"
	_LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert a Frame with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to insert a Text Frame. Error:" & @error & " Extended:" & @extended)

	;Set the Frame Style to "Labels"
	_LOWriter_FrameStyleSet($oDoc, $oFrame, "Labels")
	If (@error > 0) Then _ERROR("Failed to set the Text Frame style. Error:" & @error & " Extended:" & @extended)

	;Retrieve an Array of all available Frame Style Names.
	$asNames = _LOWriter_FrameStylesGetNames($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now insert a list of all available Frame styles. There are " & @extended & " results.")

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The available Frame Styles in this document are:" & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asNames) -1
		;Insert the Frame Style names.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asNames[$i] & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	;Retrieve an Array of all user-created Frame Style Names.
	$asNames = _LOWriter_FrameStylesGetNames($oDoc, True)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now insert a list of all user-created Frame styles. There are " & @extended & " results.")

	;Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START,1,True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the right once, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT,1,True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Frame Styles that are user-created are:" & @CR & @CR,True)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asNames) -1
		;Insert the Frame Style name.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asNames[$i] & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	;Retrieve an Array of all applied Frame Style Names.
	$asNames = _LOWriter_FrameStylesGetNames($oDoc, False, True)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now insert a list of all Frame styles used in this document. There are " & @extended & " results.")

	;Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START,1,True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Move the View Cursor to the right once, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT,1,True)
	If (@error > 0) Then _ERROR("Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Frame Styles currently in use in this document are:" & @CR & @CR,True)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asNames) -1
		;Insert the Frame Style name.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asNames[$i] & @CR)
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

