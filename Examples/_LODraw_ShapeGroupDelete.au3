#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oShape, $oTextBox, $oGroup
	Local $aoShapes[2]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Textbox.
	$oTextBox = _LODraw_ShapeTextBoxInsert($oPage, 10500, 5000, -1, 1500)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a line around the Text Box so it's visible.
	_LODraw_ShapeLineProperties($oTextBox, $LOD_SHAPE_LINE_STYLE_CONTINUOUS)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 3000 Wide by 6000 High.
	$oShape = _LODraw_DrawShapeInsert($oPage, $LOD_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 2000, 3500)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add both shapes to an array.
	$aoShapes[0] = $oTextBox
	$aoShapes[1] = $oShape

	; Create a group containing both the Textbox and rectangle.
	$oGroup = _LODraw_ShapeGroupCreate($oPage, $aoShapes)
	If @error Then _ERROR($oDoc, "Failed to create a shape group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have created a group. Press ok to ungroup the group.")

	; Ungroup the group of shapes.
	_LODraw_ShapeGroupDelete($oGroup)
	If @error Then _ERROR($oDoc, "Failed to ungroup a group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
