
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
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

	;Insert a Input Field at the View Cursor. Set the Reference to "New Input", and Text to "Some Text"
	$oField = _LOWriter_FieldFuncInputInsert($oDoc, $oViewCursor, False, "New Input", "Some Text")
	If (@error > 0) Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Input Field settings.")

	;Modify the Input Field settings. Set the Reference to "An Input", and Text to "Different Text"
	_LOWriter_FieldFuncInputModify($oField, "An Input", "Different Text")
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldFuncInputModify($oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"The Input Field's Reference is: " & $avSettings[0] & @CRLF & _
			"The Input Field's current Text is: " & $avSettings[1])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



