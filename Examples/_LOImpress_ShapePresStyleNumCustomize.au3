#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oStyle
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_TITLE
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_TITLE)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the "Outline 1" Presentation Style.
	$oStyle = _LOImpress_ShapePresStyleGetObjByname($oDoc, "outline1")
	If @error Then _ERROR($oDoc, "Failed to retrieve style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Numbering Customization settings for the Style: Modify Level 2, Numbering format = $LOI_NUM_FRMT_ARABIC, Start at 2, Color = $LO_COLOR_GOLD,
	; Relative size = 75%, Separator before = ~ , Separator after = #.
	_LOImpress_ShapePresStyleNumCustomize($oDoc, $oStyle, 2, $LOI_NUM_FRMT_ARABIC, 2, $LO_COLOR_GOLD, 75, "~", "#")
	If @error Then _ERROR($oDoc, "Failed to set Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Numbering settings for level 2. Return will be an array in order of function parameters, Return will only have
	; six elements, because Numbering Style is not set to special, so there will not be a Bullet or Char Decimal value.
	$avSettings = _LOImpress_ShapePresStyleNumCustomize($oDoc, $oStyle, 2)
	If @error Then _ERROR($oDoc, "Failed to retrieve Numbering Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Numbering style's current Customization settings for level 2 are as follows: " & @CRLF & _
			"The Number format used is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The Numbering starts at: " & $avSettings[1] & @CRLF & _
			"The color of  the numbering is (as a RGB Color Integer): " & $avSettings[2] & @CRLF & _
			"The relative size of the numbering to the font size is, as a percentage: " & $avSettings[3] & @CRLF & _
			"The Separator before the Numbering symbol is: " & $avSettings[4] & @CRLF & _
			"The Separator After the Numbering symbol is: " & $avSettings[5])

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
