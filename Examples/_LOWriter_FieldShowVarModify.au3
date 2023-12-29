
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShowVarField2, $oShowVarField3
	Local $iFormatKey
	Local $sMasterFieldName
	Local $avSettings

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
	_LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "2300", False)
	If (@error > 0) Then _ERROR("Failed to insert a text field. Error:" & @error & " Extended:" & @extended)

	;Insert some more text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert another Set Variable Field and using a different name, and allowing the function to create the MasterField, set the Value to 1260
	_LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, "ADifferentName", "1260", False)
	If (@error > 0) Then _ERROR("Failed to insert a text field. Error:" & @error & " Extended:" & @extended)

	;Insert a new line mark.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert 3 Show Variable Fields and use the new SetVar name, which is also the MasterField's name.
	_LOWriter_FieldShowVarInsert($oDoc, $oViewCursor, $sMasterFieldName, False)
	If (@error > 0) Then _ERROR("Failed to insert a text field. Error:" & @error & " Extended:" & @extended)

	;Insert a new line mark.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Make this field show the Master Field's name instead of its value.
	$oShowVarField2 = _LOWriter_FieldShowVarInsert($oDoc, $oViewCursor, $sMasterFieldName, False, Null, True)
	If (@error > 0) Then _ERROR("Failed to insert a text field. Error:" & @error & " Extended:" & @extended)

	;Insert a new line mark.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Retrieve this Number Format key, #,##0
	$iFormatKey = _LOWriter_FormatKeyCreate($oDoc, "#,##0")
	If (@error > 0) Then _ERROR("Failed to retrieve Number Format Key. Error:" & @error & " Extended:" & @extended)

	;Set Number format to the one I just retrieved.
	$oShowVarField3 = _LOWriter_FieldShowVarInsert($oDoc, $oViewCursor, $sMasterFieldName, False, $iFormatKey)
	If (@error > 0) Then _ERROR("Failed to insert a text field. Error:" & @error & " Extended:" & @extended)

	;Modify the third Show Variable Field, set the SetVar Name to "ADifferentName".
	_LOWriter_FieldShowVarModify($oDoc, $oShowVarField3, "ADifferentName")
	If (@error > 0) Then _ERROR("Failed to modify a text field. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings for the second Show Var Field. Return will be an Array with elements in order of function parameters.
	$avSettings = _LOWriter_FieldShowVarModify($oDoc, $oShowVarField2)
	If (@error > 0) Then _ERROR("Failed to retrieve text field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"The Set Variable Field name referenced is: " & $avSettings[0] & @CRLF & _
			"The Number Format Key used is: " & $avSettings[1] & @CRLF & _
			"Is the Set Variable's name displayed instead of the value? True/False: " & $avSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

