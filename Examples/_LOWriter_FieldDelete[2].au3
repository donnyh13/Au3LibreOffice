#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $iResults
	Local $sMasterFieldName
	Local $asMasters

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sMasterFieldName = "TestMaster"

	; Create a new Set Variable Master Field named "TestMaster".
	_LOWriter_FieldSetVarMasterCreate($oDoc, $sMasterFieldName)
	If @error Then _ERROR($oDoc, "Failed to create a Set Variable Master. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Set Variable Field and use the new MasterField's name., set the Value to 2300
	$oField = _LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "2300", False)
	If @error Then _ERROR($oDoc, "Failed to insert a text field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to delete the Set Variable Field.")

	; Delete the Field and its MasterField.
	_LOWriter_FieldDelete($oField, True)
	If @error Then _ERROR($oDoc, "Failed to delete a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Set Variable Master Field names.
	$asMasters = _LOWriter_FieldSetVarMastersGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of Set Variable Masters. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iResults = @extended

	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "Field Master Names currently in this Document: (Notice my newly created Master Field is not listed." & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To $iResults - 1
		; Write each Master Field name in the document.
		_LOWriter_CursorInsertString($oDoc, $oViewCursor, $asMasters[$i] & @CR)
		If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
