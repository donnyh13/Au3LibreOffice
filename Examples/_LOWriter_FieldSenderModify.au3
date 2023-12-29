
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Sender Field at the View Cursor. Fixed =True, Content = "AutoIt©", Data Type to $LOW_FIELD_USER_DATA_COMPANY
	$oField = _LOWriter_FieldSenderInsert($oDoc, $oViewCursor, False, True, "AutoIt©", $LOW_FIELD_USER_DATA_COMPANY)
	If (@error > 0) Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Sender Field.")

	;Modify the Sender Field settings. Fixed = False, Skip Content, User data = $LOW_FIELD_USER_DATA_FIRST_NAME
	_LOWriter_FieldSenderModify($oField, False, Null, $LOW_FIELD_USER_DATA_FIRST_NAME)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldSenderModify($oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"Is the Sender Field's content fixed? True/False: " & $avSettings[0] & @CRLF & _
			"The Sender Field's current content is, (This may be blank if your name is not filled in in L.O.: " & $avSettings[1] & @CRLF & _
			"The Type of Sender data to display is, (See UDF Constants): " & $avSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

