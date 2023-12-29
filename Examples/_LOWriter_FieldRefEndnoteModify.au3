
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oEndnote, $oField, $oEndnote2
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted an Endnote at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Endnote at the ViewCursor.
	$oEndnote = _LOWriter_EndnoteInsert($oDoc, $oViewCursor, False)
	If (@error > 0) Then _ERROR("Failed to insert an Endnote. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Endnote Reference Field. Reference the Endnote I inserted using its Object, do not overwrite any content selected by the cursor,
	;and refer using $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED
	$oField = _LOWriter_FieldRefEndnoteInsert($oDoc, $oViewCursor, $oEndnote, False, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED)
	If (@error > 0) Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	;Insert some more text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "I have inserted a 2nd Endnote at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Endnote at the ViewCursor.
	$oEndnote2 = _LOWriter_EndnoteInsert($oDoc, $oViewCursor, False)
	If (@error > 0) Then _ERROR("Failed to insert an Endnote. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Endnote Reference Field.")

	;Modify the Endnote Reference Field settings. Set the Referenced Endnote to Second Endnote, and Refer using $LOW_FIELD_REF_USING_ABOVE_BELOW
	_LOWriter_FieldRefEndnoteModify($oDoc, $oField, $oEndnote2, $LOW_FIELD_REF_USING_REF_TEXT)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldRefEndnoteModify($oDoc, $oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"The Endnote's Label this is being referenced is: " & _LOWriter_EndnoteModifyAnchor($avSettings[0]) & @CRLF & _
			"The Endnote is being referenced using this format, (see UDF Constants): " & $avSettings[1])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

