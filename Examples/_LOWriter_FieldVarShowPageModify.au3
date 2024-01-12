#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $iSetting

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Show Page Field at the View Cursor. Set Numbering Format to: $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N
	$oField = _LOWriter_FieldVarShowPageInsert($oDoc, $oViewCursor, False, $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N)
	If @error Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Show Page Field settings.")

	; Modify the Show Page Field settings. Set the Number format to $LOW_NUM_STYLE_ROMAN_UPPER
	_LOWriter_FieldVarShowPageModify($oField, $LOW_NUM_STYLE_ROMAN_UPPER)
	If @error Then _ERROR("Failed to modify field. Error:" & @error & " Extended:" & @extended)

	; Retrieve current Field settings. Return will be an integer.
	$iSetting = _LOWriter_FieldVarShowPageModify($oField)
	If @error Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Show Page Field number format setting is, (see UDF Constants): " & $iSetting)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
