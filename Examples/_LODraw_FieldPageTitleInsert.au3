#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oTextBox, $oTextCursor

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add another Page.
	$oPage = _LODraw_PageAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to insert a new page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the current Page to Page 2.
	_LODraw_PageCurrent($oDoc, $oPage)
	If @error Then _ERROR($oDoc, "Failed to set the current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a second Page.
	_LODraw_PageAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to insert a new page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Text Box.
	$oTextBox = _LODraw_ShapeTextBoxInsert($oPage, 15000, 13000)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a TextCursor in the TextBox.
	$oTextCursor = _LODraw_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a TextCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Page Title field in the TextBox.
	_LODraw_FieldPageTitleInsert($oDoc, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
