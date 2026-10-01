#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oTextBox, $oTextCursor
	Local $sString
	Local $avFields
	Local $iResults

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Text Box.
	$oTextBox = _LODraw_ShapeTextBoxInsert($oPage, 15000, 13000)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a TextCursor in the TextBox.
	$oTextCursor = _LODraw_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a TextCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Date field in the TextBox.
	_LODraw_FieldDateTimeInsert($oDoc, $oTextCursor, True, False, Null, $LOD_FIELD_DATE_FMT_DOW_MMMM_DD_YYYY)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a space between fields.
	_LODraw_CursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert an Author field in the TextBox, Set it to be Fixed, and set the Author to "AutoIt".
	_LODraw_FieldAuthorInsert($oDoc, $oTextCursor, True, "AutoIt", $LOD_FIELD_AUTH_NAME_FIRST)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a space between fields.
	_LODraw_CursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a File Name field in the TextBox, Set it to be variable, and File Name format to just the name.
	_LODraw_FieldFileNameInsert($oDoc, $oTextCursor, False, $LOD_FIELD_FILENAME_NAME)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a space between fields.
	_LODraw_CursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Hyperlink field in the TextBox.
	_LODraw_FieldHyperlinkInsert($oDoc, $oTextCursor, "www.autoitscript.com/site/autoit/", "AutoIt v3")
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a space between fields.
	_LODraw_CursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Page Count field in the TextBox.
	_LODraw_FieldPageCountInsert($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a space between fields.
	_LODraw_CursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Page Number field in the TextBox.
	_LODraw_FieldPageNumberInsert($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a space between fields.
	_LODraw_CursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Page Title field in the TextBox.
	_LODraw_FieldPageTitleInsert($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of all Field Objects.
	$avFields = _LODraw_FieldsGetList($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of fields. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iResults = @extended

	For $i = 0 To UBound($avFields) - 1
		$sString &= "This is the Field's Object: (IsObj: " & IsObj($avFields[$i][0]) & ")" & @CRLF & _
				"The Field Constant number is (See UDF Constants): " & $avFields[$i][1] & @CRLF & _
				"And the current display of the field is: " & _LODraw_FieldCurrentDisplayGet($avFields[$i][0]) & @CRLF & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I Found " & $iResults & " fields, the Fields found are: " & @CRLF & @CRLF & $sString)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
