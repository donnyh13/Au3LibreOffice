#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oTextCursor, $oCell, $oAnchorCursor
	Local $mField

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell A1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Create a Text Cursor in the cell.
	$oTextCursor = _LOCalc_CellCreateTextCursor($oCell)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in Cell. Error:" & @error & " Extended:" & @extended)

	; Insert a URL field in the cell A1.
	$mField = _LOCalc_FieldHyperlinkInsert($oDoc, $oTextCursor, "https://www.autoitscript.com/site/autoit/", "AutoIt")
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Create a Text Cursor at the Field's Anchor.
	$oAnchorCursor = _LOCalc_FieldGetAnchor($mField)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor at the Field's Anchor. Error:" & @error & " Extended:" & @extended)

	; Move the cursor to the beginning of the text.
	_LOCalc_TextCursorMove($oAnchorCursor, $LOC_TEXTCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Failed to move Text Cursor. Error:" & @error & " Extended:" & @extended)

	; Insert some text before the anchor.
	_LOCalc_TextCursorInsertString($oAnchorCursor, "This UDF was created with ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
