
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $iTimeFormatKey
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

	;Create or retrieve a DateFormat Key, Hour, Minute, Second
	$iTimeFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "HH:MM:SS")
	If (@error > 0) Then _ERROR("Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended)

	;Set the Document's General Property settings, Set total editing time to 5137 seconds (1:25:37)
	_LOWriter_DocGenProp($oDoc, Null, Null, 5137)
	If (@error > 0) Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	;Insert a Doc Info Editing Time Field at the View Cursor. Set is Fixed = True, Time format key to the one I just created.
	$oField = _LOWriter_FieldDocInfoEditTimeInsert($oDoc, $oViewCursor, False, True, $iTimeFormatKey)
	If (@error > 0) Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Doc Info Field settings.")

	;Create or retrieve a different DateFormat Key, Hour, Minute
	$iTimeFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "H:MM:")
	If (@error > 0) Then _ERROR("Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended)

	;Modify the Doc Info Editing time Field settings. Set Fixed to False, and set the new Time Format key to be used.
	_LOWriter_FieldDocInfoEditTimeModify($oDoc, $oField, False, $iTimeFormatKey)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an array, with elements in order of function parameters.
	$avSettings = _LOWriter_FieldDocInfoEditTimeModify($oDoc, $oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Doc Info Field settings are: " & @CRLF & _
			"Is the content of this field fixed? True/ False: " & $avSettings[0] & @CRLF & _
			"The Time Format key used is: " & $avSettings[1] & " and looks like: " & @CRLF & _
			_LOWriter_DateFormatKeyGetString($oDoc, $avSettings[1]))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

