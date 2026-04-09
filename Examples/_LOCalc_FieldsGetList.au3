#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oTextCursor, $oCell
	Local $avFields[0][0]
	Local $sString = ""
	Local $iResults = 0

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the cell.
	$oTextCursor = _LOCalc_CellCreateTextCursor($oCell)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in Cell. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Document Title field in the cell A1.
	_LOCalc_FieldTitleInsert($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor 4. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	;Insert a space after the field.
	_LOCalc_TextCursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert a String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Sheet name field in the cell A1.
	_LOCalc_FieldSheetNameInsert($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor 2. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	;Insert a space after the field.
	_LOCalc_TextCursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert a String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Date field in the cell A1.
	_LOCalc_FieldDateTimeInsert($oDoc, $oTextCursor, True)
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor 3. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	;Insert a space after the field.
	_LOCalc_TextCursorInsertString($oTextCursor, " ")
	If @error Then _ERROR($oDoc, "Failed to insert a String. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a URL field in the cell A1.
	_LOCalc_FieldHyperlinkInsert($oDoc, $oTextCursor, "https://www.autoitscript.com/site/autoit/", "AutoIt")
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor 1. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of all Field Objects.
	$avFields = _LOCalc_FieldsGetList($oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of fields. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iResults = @extended

	For $i = 0 To UBound($avFields) - 1
		$sString &= "The Field's command name is: " & _LOCalc_FieldCurrentDisplayGet($avFields[$i][0], True) & @CRLF & _
				"The Field Constant number is: " & $avFields[$i][1] & @CRLF & _
				"And the current display of the field is: " & _LOCalc_FieldCurrentDisplayGet($avFields[$i][0]) & @CRLF & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I Found " & $iResults & " fields, the Fields found are: " & @CRLF & @CRLF & $sString)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
