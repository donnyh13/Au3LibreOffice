#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $avArray
	Local $iNewX, $iNewY, $iCount

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a FreeForm Line Shape into the document, 5000 Wide by 0 High, 12000X, 4300Y.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE, 5000, 0, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current number of points in the shape.
	$iCount = _LOImpress_DrawShapePointsGetCount($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of Points contained in the shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The number of points in this shape is: " & $iCount)

	; Retrieve the Shape's current settings for its first point.
	$avArray = _LOImpress_DrawShapePointsModify($oShape, 1)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of settings for a Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; I will retrieve the first point's current position, and add to its X and Y values to determine my new point's new X and Y values.

	; Minus 400 Hundredths of a Millimeter (100th MM) from the X coordinate
	$iNewX = $avArray[0] - 400

	; Add 1600 Hundredths of a Millimeter (100th MM) to the Y coordinate
	$iNewY = $avArray[1] + 1600

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to insert the new Point into the shape.")

	; Add the new Point using the new X and Y coordinates. The new point will be added after the called point (1), The point's type will be Symmetrical.
	_LOImpress_DrawShapePointsAdd($oShape, 1, $iNewX, $iNewY, $LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC)
	If @error Then _ERROR($oDoc, "Failed to modify Shape point. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the new number of points in the shape.
	$iCount = _LOImpress_DrawShapePointsGetCount($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve a count of Points contained in the shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The new number of points in this shape is: " & $iCount)

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
