#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Placeholder Field at the View Cursor. Set the Placeholder type to $LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC, Name = "A Graphic Placeholder.",
	; Reference = "Click Me"
	$oField = _LOWriter_FieldFuncPlaceholderInsert($oDoc, $oViewCursor, False, $LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC, "A Graphic Placeholder.", "Click Me")
	If @error Then _ERROR($oDoc, "Failed to insert a Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Placeholder Field settings.")

	; Modify the Placeholder Field settings. Set the Placeholder type to $LOW_FIELD_PLACEHOLD_TYPE_TABLE, Set the Name to "A Table Placeholder",
	; And Reference to "Hover Me"
	_LOWriter_FieldFuncPlaceholderModify($oField, $LOW_FIELD_PLACEHOLD_TYPE_TABLE, "A Table Placeholder", "Hover Me")
	If @error Then _ERROR($oDoc, "Failed to modify field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldFuncPlaceholderModify($oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Field settings are: " & @CRLF & _
			"The Placeholder Field Type is, (see UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Placeholder Field name is: " & $avSettings[1] & @CRLF & _
			"The Placeholder Field's Reference Text is: " & $avSettings[2])

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
