
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iTimeFormatKey

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
	_LOWriter_FieldDocInfoEditTimeInsert($oDoc, $oViewCursor, False, True, $iTimeFormatKey)
	If (@error > 0) Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



