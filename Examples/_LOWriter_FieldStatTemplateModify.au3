
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $iSetting

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Set the Document's template settings for use in this demonstration to, Template name = "AutoIt", Template URL to a fake Path, @TempDir\Folder2\AutoIt.ott
	_LOWriter_DocGenPropTemplate($oDoc, "AutoIt", @TempDir & "\Folder2\AutoIt.ott")
	If (@error > 0) Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Show Template at the View Cursor. Set display format to $LOW_FIELD_FILENAME_FULL_PATH
	$oField = _LOWriter_FieldStatTemplateInsert($oDoc, $oViewCursor, False, $LOW_FIELD_FILENAME_FULL_PATH)
	If (@error > 0) Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Template Field settings.")

	;Modify the Template Field settings. Set the display format to $LOW_FIELD_FILENAME_NAME_AND_EXT
	_LOWriter_FieldStatTemplateModify($oField, $LOW_FIELD_FILENAME_NAME_AND_EXT)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an integer.
	$iSetting = _LOWriter_FieldStatTemplateModify($oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Template Field display format setting is, (see UDF Constants): " & $iSetting)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



