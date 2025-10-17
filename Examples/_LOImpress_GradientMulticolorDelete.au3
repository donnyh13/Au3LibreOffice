#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $avStops
	Local $sStops = ""

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the document, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the Shape.
	_LOImpress_DrawShapePosition($oShape, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to move a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Gradient settings to: Preset Gradient name = $LOI_GRAD_NAME_SUNDOWN
	_LOImpress_DrawShapeAreaGradient($oShape, $LOI_GRAD_NAME_SUNDOWN)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Multicolor Gradient ColorStops.
	$avStops = _LOImpress_DrawShapeAreaGradientMulticolor($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a new ColorStop in the middle.
	_LOImpress_GradientMulticolorAdd($avStops, 3, 0.6, 1234567)
	If @error Then _ERROR($oDoc, "Failed to add a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add another new ColorStop in the middle.
	_LOImpress_GradientMulticolorAdd($avStops, 5, 0.8, 654321)
	If @error Then _ERROR($oDoc, "Failed to add a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new ColorStops.
	_LOImpress_DrawShapeAreaGradientMulticolor($oShape, $avStops)
	If @error Then _ERROR($oDoc, "Failed to modify Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Multicolor Gradient ColorStops.
	$avStops = _LOImpress_DrawShapeAreaGradientMulticolor($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Color: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Gradient ColorStops are as follows: " & @CRLF & _
			$sStops & @CRLF & "Press ok to delete the first and last ColorStop.")

	; Delete the first ColorStop
	_LOImpress_GradientMulticolorDelete($avStops, 0)
	If @error Then _ERROR($oDoc, "Failed to delete a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the last ColorStop
	_LOImpress_GradientMulticolorDelete($avStops, (UBound($avStops) - 1))
	If @error Then _ERROR($oDoc, "Failed to delete a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new ColorStops.
	_LOImpress_DrawShapeAreaGradientMulticolor($oShape, $avStops)
	If @error Then _ERROR($oDoc, "Failed to modify Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Multicolor Gradient ColorStops.
	$avStops = _LOImpress_DrawShapeAreaGradientMulticolor($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStops = ""

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Color: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now the Shape's Gradient ColorStops are as follows: " & @CRLF & _
			$sStops)

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
