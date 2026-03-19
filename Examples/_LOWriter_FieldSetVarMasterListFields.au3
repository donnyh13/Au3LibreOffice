
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oMasterfield, $oViewCursor
	Local $iResults
	Local $sMasterFieldName
	Local $aoFields

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	$sMasterFieldName = "TestMaster"

	;Create a new Set Variable Master Field named "TestMaster".
	$oMasterfield = _LOWriter_FieldSetVarMasterCreate($oDoc, $sMasterFieldName)
	If (@error > 0) Then _ERROR("Failed to create a Set Variable Master. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert a Set Variable Field and use the new MasterField's name., set the Value to 2300
	_LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "2300", False)
	If (@error > 0) Then _ERROR("Failed to insert a text field. Error:" & @error & " Extended:" & @extended)

	;Retrieve an array of dependent fields for the Master Field. Dependent fields are Set Variable fields that reference the Master.
	$aoFields = _LOWriter_FieldSetVarMasterListFields($oDoc, $oMasterfield)
	If (@error > 0) Then _ERROR("Failed to retrieve an array of Dependent fields. Error:" & @error & " Extended:" & @extended)

	$iResults = @extended

	MsgBox($MB_OK, "", "I found " & $iResults & " dependent fields for this Master Field. Press Ok to delete one of these fields.")

	;Delete the last Field result.
	_LOWriter_FieldDelete($oDoc, $aoFields[$iResults - 1])
	If (@error > 0) Then _ERROR("Failed to delete a Dependent field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



