#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oShape
	Local $sStops = ""
	Local $avStops

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Circle Shape into the document, 5000 Wide by 5000 High.
	$oShape = _LOWriter_ShapeInsert($oDoc, $oViewCursor, $LOW_SHAPE_TYPE_BASIC_CIRCLE, 5000, 5000)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape  Background Color settings. Background color = $LO_COLOR_TEAL.
	_LOWriter_ShapeAreaColor($oShape, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Transparency Gradient settings to: Gradient Type = $LOW_GRAD_TYPE_RADIAL, XCenter to 50%, YCenter to 50%, Angle to 0 degrees
	; Border to 0%, Start transparency to 100%, End Transparency to 0%
	_LOWriter_ShapeAreaTransparencyGradient($oDoc, $oShape, $LOW_GRAD_TYPE_RADIAL, 50, 50, 0, 0, 100, 0)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Transparency Multi Gradient ColorStops.
	$avStops = _LOWriter_ShapeAreaTransparencyGradientMulti($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a new ColorStop in the middle.
	_LOWriter_TransparencyGradientMultiAdd($avStops, 1, 0.4, 55)
	If @error Then _ERROR($oDoc, "Failed to add a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add another new ColorStop in the middle.
	_LOWriter_TransparencyGradientMultiAdd($avStops, 2, 0.7, 70)
	If @error Then _ERROR($oDoc, "Failed to add a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new ColorStops.
	_LOWriter_ShapeAreaTransparencyGradientMulti($oShape, $avStops)
	If @error Then _ERROR($oDoc, "Failed to modify Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Transparency Multi Gradient ColorStops.
	$avStops = _LOWriter_ShapeAreaTransparencyGradientMulti($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Transparency percentage: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Transparency Gradient ColorStops are as follows: " & @CRLF & _
			$sStops & @CRLF & "Press ok to modify the first and last ColorStop.")

	; Modify the first ColorStop
	_LOWriter_TransparencyGradientMultiModify($avStops, 0, 0.2, 10)
	If @error Then _ERROR($oDoc, "Failed to modify a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the last ColorStop
	_LOWriter_TransparencyGradientMultiModify($avStops, (UBound($avStops) - 1), 0.9, 85)
	If @error Then _ERROR($oDoc, "Failed to modify a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new ColorStops.
	_LOWriter_ShapeAreaTransparencyGradientMulti($oShape, $avStops)
	If @error Then _ERROR($oDoc, "Failed to modify Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Transparency Multi Gradient ColorStops.
	$avStops = _LOWriter_ShapeAreaTransparencyGradientMulti($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Transparency Multi Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStops = ""

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Transparency percentage: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now the Shape's Transparency Gradient ColorStops are as follows: " & @CRLF & _
			$sStops)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
