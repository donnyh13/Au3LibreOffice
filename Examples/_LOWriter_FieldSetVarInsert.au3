#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iFormatKey
	Local $sMasterFieldName

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
	_LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "2300", False)
	If @error Then _ERROR($oDoc, "Failed to insert a text field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert 2 new lines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve this Number Format key, #,##0
	$iFormatKey = _LOWriter_FormatKeyCreate($oDoc, "#,##0")
	If @error Then _ERROR($oDoc, "Failed to retrieve Number Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Set Variable Field and use a new name, and allow the function to create a Master Field for me., set the Value to 1260, Number format key =
	; #,##0, Visible = True.
	_LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "1260", False, $iFormatKey, True)
	If @error Then _ERROR($oDoc, "Failed to insert a text field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
