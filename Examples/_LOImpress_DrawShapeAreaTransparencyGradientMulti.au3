#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $avStops
	Local $sStops

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Rectangle Shape into the Slide, 3000 Wide by 6000 High.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, 3000, 6000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Background Color settings. Background color = $LO_COLOR_TEAL.
	_LOImpress_DrawShapeAreaColor($oShape, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Transparency Gradient settings to: Gradient Type = $LOI_GRAD_TYPE_ELLIPTICAL, XCenter to 75%, YCenter to 45%, Angle to 180 degrees
	; Border to 16%, Start transparency to 10%, End Transparency to 62%
	_LOImpress_DrawShapeAreaTransparencyGradient($oShape, $LOI_GRAD_TYPE_ELLIPTICAL, 75, 45, 180, 16, 10, 62)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Transparency Multi Gradient ColorStops.
	$avStops = _LOImpress_DrawShapeAreaTransparencyGradientMulti($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Transparency percentage: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Transparency Gradient ColorStops are as follows: " & @CRLF & _
			$sStops & @CRLF & @CRLF & _
			"Press ok to add a new ColorStop.")

	; Add a new ColorStop in the middle.
	_LOImpress_TransparencyGradientMultiAdd($avStops, 1, 0.5, 76)
	If @error Then _ERROR($oDoc, "Failed to add a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new ColorStops.
	_LOImpress_DrawShapeAreaTransparencyGradientMulti($oShape, $avStops)
	If @error Then _ERROR($oDoc, "Failed to modify Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Transparency Multi Gradient ColorStops.
	$avStops = _LOImpress_DrawShapeAreaTransparencyGradientMulti($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStops = ""

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Transparency percentage: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Transparency Gradient ColorStops are as follows: " & @CRLF & _
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
