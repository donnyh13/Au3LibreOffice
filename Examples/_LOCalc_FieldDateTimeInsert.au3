#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor, $oPageStyle, $oHeader
	Local $sPageStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell A1
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A1")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the cell.
	$oTextCursor = _LOCalc_CellCreateTextCursor($oCell)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in Cell. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Date Field in the cell.
	_LOCalc_FieldDateTimeInsert($oDoc, $oTextCursor, True)
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the currently active Sheet's Page Style name.
	$sPageStyle = _LOCalc_PageStyleSet($oDoc, $oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Page Style object.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, $sPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style object by name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Header Object
	$oHeader = _LOCalc_PageStyleHeaderObj($oPageStyle, Default)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style header object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the First Page Header, left side.
	$oTextCursor = _LOCalc_PageStyleHeaderCreateTextCursor($oHeader, True, True)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in header. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Time Field in the Header.
	_LOCalc_FieldDateTimeInsert($oDoc, $oTextCursor, False)
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor 2. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the modified Header Object
	_LOCalc_PageStyleHeaderObj($oPageStyle, $oHeader)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style header object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK, "", "I have inserted a date field in cell A1, and also a Time Field on the left side first page header." & @CRLF & _
			"To see the field in Left side header, go to: Format->Page Style->Header->Edit. The Field will be in ""Left Area""." & @CRLF & _
			"When finished, please close the opened pages, or else the document will not close correctly.")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
