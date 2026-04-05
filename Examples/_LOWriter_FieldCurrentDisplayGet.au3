#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oCondTextField, $oAuthorField, $oSetVarField, $oShowVarField
	Local $sMasterFieldName

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_ViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Conditional text Field at the View Cursor. The Condition will be if (1+1) > [is greater than] 2 If so display "Yes", If not, Display "NO!"
	$oCondTextField = _LOWriter_FieldCondTextInsert($oDoc, $oViewCursor, False, "(1+1) > 2", "Yes", "No!")
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Author Field at the View Cursor. Fixed = True, Set the Author to "Auto It", Full name = True.
	$oAuthorField = _LOWriter_FieldAuthorInsert($oDoc, $oViewCursor, False, True, "Auto It", True)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sMasterFieldName = "TestMaster"

	; Create a new Set Variable Master Field named "TestMaster".
	_LOWriter_FieldSetVarMasterCreate($oDoc, $sMasterFieldName)
	If @error Then _ERROR($oDoc, "Failed to create a Set Variable Master. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Set Variable Field and use the new MasterField's name, set the Value to 2300
	$oSetVarField = _LOWriter_FieldSetVarInsert($oDoc, $oViewCursor, $sMasterFieldName, "2300", False)
	If @error Then _ERROR($oDoc, "Failed to insert a text field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Show Variable Field and use the new SetVar name, which is also the MasterField's name.
	$oShowVarField = _LOWriter_FieldShowVarInsert($oDoc, $oViewCursor, $sMasterFieldName, False)
	If @error Then _ERROR($oDoc, "Failed to insert a text field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Current displayed text of the Conditional Text Field is: " & _LOWriter_FieldCurrentDisplayGet($oCondTextField) & @CRLF & _
			"The Current displayed text of the Author Field is: " & _LOWriter_FieldCurrentDisplayGet($oAuthorField) & @CRLF & _
			"The Current displayed text of the Set Variable Field is: " & _LOWriter_FieldCurrentDisplayGet($oSetVarField) & @CRLF & _
			"The Current displayed text of the Show Variable Field is: " & _LOWriter_FieldCurrentDisplayGet($oShowVarField))

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
