
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oMaster
	Local $sMasterFieldName

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	$sMasterFieldName = "TestMaster"

	;Create a new Set Variable Master Field named "TestMaster".
	$oMaster = _LOWriter_FieldSetVarMasterCreate($oDoc, $sMasterFieldName)
	If (@error > 0) Then _ERROR("Failed to create a Set Variable Master. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does the Set Var Master Field exist in the Document? True/False: " & _LOWriter_FieldSetVarMasterExists($oDoc, $sMasterFieldName))

	MsgBox($MB_OK, "", "Press ok to delete the newly created Set Variable Master Field.")

	;Delete the Set Var. MasterField.
	_LOWriter_FieldSetVarMasterDelete($oDoc, $oMaster)
	If (@error > 0) Then _ERROR("Failed to delete Master Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does the Set Var Master Field still exist? True/False: " & _LOWriter_FieldSetVarMasterExists($oDoc, $sMasterFieldName))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

