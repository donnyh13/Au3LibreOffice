#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"


Example()

Func Example()
	Local $oDoc, $oSlide, $oShape, $oTextBox, $oGroup, $oGroupedShape
	Local $iType
	Local $aoShapes[2]

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active Slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_BLANK
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_BLANK)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Textbox.
	$oTextBox = _LOImpress_ShapeTextBoxInsert($oSlide, $LOI_SHAPE_TEXTBOX_TYPE_TEXTBOX, 10500, 5000, -1, 1500)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a line around the Text Box so it's visible.
	_LOImpress_ShapeLineProperties($oTextBox, $LOI_SHAPE_LINE_STYLE_CONTINUOUS)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Page, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 2000, 3500)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add both shapes to an array.
	$aoShapes[0] = $oTextBox
	$aoShapes[1] = $oShape

	; Create a group containing both the Textbox and rectangle.
	$oGroup = _LOImpress_ShapeGroupCreate($oSlide, $aoShapes)
	If @error Then _ERROR($oDoc, "Failed to create a shape group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Shape's Object by Index
	$oGroupedShape = _LOImpress_ShapeGroupShapeGetObjByIndex($oGroup, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape Object from Group. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iType = @extended

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have retrieved the Object for the 0th shape in the group, which has a shape type of (See UDF Constants): " & $iType & @CRLF & _
			"Press ok to change its background color.")

	; Modify the Shape Background Color settings. Background color = $LO_COLOR_BROWN.
	_LOImpress_ShapeAreaColor($oGroupedShape, $LO_COLOR_BROWN)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
