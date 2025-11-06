#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oShape
	Local $i100thMM
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Folded Corner Shape into the document, 3000 Wide by 6000 High, 12000X, 4300Y.
	$oShape = _LOImpress_DrawShapeInsert($oSlide, $LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER, 3000, 6000, 12000, 4300)
	If @error Then _ERROR($oDoc, "Failed to create a Shape. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/8" to Hundredths of a Millimeter (100th MM)
	$i100thMM = _LO_UnitConvert(.125, $LO_CONVERT_UNIT_INCH_100THMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (100th MM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Shape Line Properties settings to: Set the Line Style to $LOI_SHAPE_LINE_STYLE_3_DASHES_3_DOTS, Line Color to $LO_COLOR_MAGENTA,
	; Width = 1/8", Transparency = 50%, Corner Style = $LOI_SHAPE_LINE_JOINT_BEVEL, Cap Style = $LOI_SHAPE_LINE_CAP_SQUARE
	_LOImpress_DrawShapeLineProperties($oShape, $LOI_DRAWSHAPE_LINE_STYLE_3_DASHES_3_DOTS, $LO_COLOR_MAGENTA, $i100thMM, 50, $LOI_DRAWSHAPE_LINE_JOINT_BEVEL, $LOI_DRAWSHAPE_LINE_CAP_SQUARE)
	If @error Then _ERROR($oDoc, "Failed to set Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_DrawShapeLineProperties($oShape)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Shape's Line Properties settings are as follows: " & @CRLF & _
			"The Line Style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Line color is, is long Color format: " & $avSettings[1] & @CRLF & _
			"The Line's Width is, in Hundredths of a Millimeter (100th MM): " & $avSettings[2] & @CRLF & _
			"The Line's transparency percentage is: " & $avSettings[3] & @CRLF & _
			"The Line Corner Style is, (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The Line Cap Style is, (See UDF Constants): " & $avSettings[5])

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
