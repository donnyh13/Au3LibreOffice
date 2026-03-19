
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

	;Insert a Doc Info Subject Field at the View Cursor. Set is Fixed = True, Subject = "This is a Subject Field."
	$oField = _LOWriter_FieldDocInfoSubjectInsert($oDoc, $oViewCursor, False, True, "This is a Subject Field.")
	If (@error > 0) Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Doc Info Field settings.")

	;Set the Document's Description settings, Set the subject to, "Autoit Automation."
	_LOWriter_DocDescription($oDoc, Null, "Autoit Automation.")
	If (@error > 0) Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	;Modify the Doc Info Subject Field settings. Set Fixed to False.
	_LOWriter_FieldDocInfoSubjectModify($oField, False)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an array, with elements in order of function parameters.
	$avSettings = _LOWriter_FieldDocInfoSubjectModify($oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Doc Info Field settings are: " & @CRLF & _
			"Is the content of this field fixed? True/ False: " & $avSettings[0] & @CRLF & _
			"The Subject is: " & $avSettings[1])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



