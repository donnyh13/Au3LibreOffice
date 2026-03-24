#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oSetVarField
	Local $iFormatKey
	Local $sMasterFieldName
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
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
	$oSetVarField = _LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "2300", False)
	If @error Then _ERROR($oDoc, "Failed to insert a text field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Set Variable Field.")

	; Retrieve this Number Format key, #,##0
	$iFormatKey = _LOWriter_FormatKeyCreate($oDoc, "#,##0")

	; Modify the Set Variable Field settings, Change the Value to 1260, Set the Number format Key to the one just retrieved
	_LOWriter_FieldSetVarModify($oDoc, $oSetVarField, "1260", $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to modify Text Field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Set Variable Field's settings.
	$avSettings = _LOWriter_FieldSetVarModify($oDoc, $oSetVarField)
	If @error Then _ERROR($oDoc, "Failed to retrieve Text Field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Field settings are: " & @CRLF & _
			"The Set Variable Field's current value is: " & $avSettings[0] & @CRLF & _
			"The Number Format key used to display the Set Variable Field's value is: " & $avSettings[1] & @CRLF & _
			"Is the Set Variable Field visible in the document? True/False: " & $avSettings[2])

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
