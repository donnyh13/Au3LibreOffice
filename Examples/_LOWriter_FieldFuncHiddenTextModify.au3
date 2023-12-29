
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $sNewCondition, $sNewText
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

	;Insert a Hidden Text Field at the View Cursor. Set the Condition to "(1+1) == 3"
	$oField = _LOWriter_FieldFuncHiddenTextInsert($oDoc, $oViewCursor, False, "(1+1) == 3", "Some Hidden Text")
	If (@error > 0) Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & "If the above Text is still visible, in your toolbar go to Tools, Options, LibreOffice Writer," & _
			" view, and under ""Display Fields"" heading, uncheck:""Hidden Text""")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	$sNewCondition = InputBox("Modify the Field.", "Enter a new condition to set the field to.", "(2*2+4) == 8")

	$sNewText = InputBox("Modify the Field.", "Enter new text to display if the condition is True.", "Some Different TEXT")

	;Modify the Hidden Text Field settings. Set the condition to the user set condition, and the Text to the New User Text.
	_LOWriter_FieldFuncHiddenTextModify($oField, $sNewCondition, $sNewText)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldFuncHiddenTextModify($oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"The Hidden Text Field's condition to evaluate is: " & $avSettings[0] & @CRLF & _
			"The Hidden Text Field's text to display if the condition is true, is: " & $avSettings[1] & @CRLF & _
			"Is the Text Hidden? True/False: " & $avSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

