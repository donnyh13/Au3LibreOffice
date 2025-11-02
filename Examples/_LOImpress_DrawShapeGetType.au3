#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $avShapes
	Local $sShapeTypes = ""

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Slide, 3000 Wide by 6000 High, 500X, 100Y.
	_LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000, 500, 100)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Connector Shape into the Slide, 1000 Wide by 3000 High, 12100X, 4400Y.
	_LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_CONNECTOR_ARROWS, 900, 3000, 12100, 4400)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Line Shape into the Slide, 4000 Wide by 0 High, 12000X, 4300Y.
	_LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW, 4000, 0, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all drawing Shapes in the current slide.
	$avShapes = _LOImpress_SlideShapesGetList($oSlide, $LOI_SHAPE_TYPE_DRAWING_SHAPE)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shapes in slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To @extended - 1
		; Get the Shape type.
		$sShapeTypes &= "Shape #" & $i + 1 & " is the following type: " & _LOImpress_DrawShapeGetType($avShapes[$i][0]) & @CRLF
		If @error Then _ERROR($oDoc, "Failed to identify shape type. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The shapes in this slide are the following types (See UDF Constants)." & @CRLF & $sShapeTypes)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
