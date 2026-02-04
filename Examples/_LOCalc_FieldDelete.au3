#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell, $oTextCursor, $oPageStyle, $oHeader
	Local $mField
	Local $sPageStyle
	Local $avFields[0]

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
	$mField = _LOCalc_FieldDateTimeInsert($oDoc, $oTextCursor, True)
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to delete the newly inserted Field from cell A1.")

	; Delete the Field
	_LOCalc_FieldDelete($mField)
	If @error Then _ERROR($oDoc, "Failed to delete field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a field into the header, and then demonstrate how to delete it.")

	; Retrieve the currently active Sheet's Page Style name.
	$sPageStyle = _LOCalc_PageStyleCurrent($oDoc, $oSheet)
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

	; Insert a Page Count field in the Header.
	_LOCalc_FieldPageCountInsert($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to insert field at Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the modified Header Object
	_LOCalc_PageStyleHeaderObj($oPageStyle, $oHeader)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style header object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have inserted the Field into the Left Side header, I will now demonstrate deleting it.")

	; Retrieve the Header Object
	$oHeader = _LOCalc_PageStyleHeaderObj($oPageStyle, Default)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style header object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the First Page Header, left side.
	$oTextCursor = _LOCalc_PageStyleHeaderCreateTextCursor($oHeader, True, True)
	If @error Then _ERROR($oDoc, "Failed to create Text Cursor in header. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Page Count fields present in the Header.
	$avFields = _LOCalc_FieldsGetList($oTextCursor, $LOC_FIELD_TYPE_PAGE_COUNT, False)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of fields in header. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the first found field. (This will be the only field).
	_LOCalc_FieldDelete($avFields[0])
	If @error Then _ERROR($oDoc, "Failed to delete Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the modified Header Object
	_LOCalc_PageStyleHeaderObj($oPageStyle, $oHeader)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style header object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
