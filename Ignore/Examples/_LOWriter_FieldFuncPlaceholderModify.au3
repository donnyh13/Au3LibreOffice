#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Placeholder Field at the View Cursor. Set the PlaceHolder type to $LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC, Name = "A Graphic PlaceHolder.",
	; Reference = "Click Me"
	$oField = _LOWriter_FieldFuncPlaceholderInsert($oDoc, $oViewCursor, False, $LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC, "A Graphic PlaceHolder.", "Click Me")
	If @error Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the PlaceHolder Field settings.")

	; Modify the PlaceHolder Field settings. Set the PlaceHolder type to $LOW_FIELD_PLACEHOLD_TYPE_TABLE, Set the Name to "A Table PlaceHolder",
	; And Reference to "Hover Me"
	_LOWriter_FieldFuncPlaceholderModify($oField, $LOW_FIELD_PLACEHOLD_TYPE_TABLE, "A Table PlaceHolder", "Hover Me")
	If @error Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	; Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldFuncPlaceholderModify($oField)
	If @error Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"The PlaceHolder Field Type is, (see UDF Constants): " & $avSettings[0] & @CRLF & _
			"The PlaceHolder Field name is: " & $avSettings[1] & @CRLF & _
			"The PlaceHolder Field's Reference Text is: " & $avSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
