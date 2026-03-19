
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $iResults
	Local $sMasterFieldName
	Local $asMasters

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	$sMasterFieldName = "TestMaster"

	;Create a new Set Variable Master Field named "TestMaster".
	_LOWriter_FieldSetVarMasterCreate($oDoc, $sMasterFieldName)
	If (@error > 0) Then _ERROR("Failed to create a Set Variable Master. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Set Variable Field and use the new MasterField's name., set the Value to 2300
	$oField = _LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "2300", False)
	If (@error > 0) Then _ERROR("Failed to insert a text field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to delete the Set Variable Field.")

	;Delete the Field and It's MasterField.
	_LOWriter_FieldDelete($oDoc, $oField, True)
	If (@error > 0) Then _ERROR("Failed to delete a field. Error:" & @error & " Extended:" & @extended)

	;Retrieve an array of Set Variable Master Field names.
	$asMasters = _LOWriter_FieldSetVarMasterList($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve an array of Set Variable Masters. Error:" & @error & " Extended:" & @extended)
	$iResults = @extended

	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "Field Master Names currently in this Document: (Notice my newly created Master Field is not listed." & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To $iResults - 1
		;Write each Master Field name in the document.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asMasters[$i] & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



