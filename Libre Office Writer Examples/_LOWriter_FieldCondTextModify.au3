
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

	;Insert a Conditional text Field at the View Cursor. The Condition will be if (1+1) > [is greater than] 2 If so display "Yes", If not, Display "NO!"
	$oField = _LOWriter_FieldCondTextInsert($oDoc, $oViewCursor, False, "(1+1) > 2", "Yes", "No!")
	If (@error > 0) Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Conditional Text Field.")

	;Modify the Conditional Text Field settings. Set the New condition to If the Page Count = 1 Then display "There is 1 Page" , else display "There are many pages"
	_LOWriter_FieldCondTextModify($oField, "PAGE == 1", "There is 1 Page", "There are many pages")
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings.
	$avSettings = _LOWriter_FieldCondTextModify($oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"The Conditional Text Field's Condition to evaluate is: " & $avSettings[0] & @CRLF & _
			"If the condition is true, display this text: " & $avSettings[1] & @CRLF & _
			"If the condition is False, display this text: " & $avSettings[2] & @CRLF & _
			"Is the condition currently evaluated as True? True/False: " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



