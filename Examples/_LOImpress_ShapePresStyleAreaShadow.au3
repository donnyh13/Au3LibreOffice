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

	; Get the Object for the "Title" Presentation Style.
	$oStyle = _LOImpress_ShapePresStyleGetObjByName($oDoc, "title")
	If @error Then _ERROR($oDoc, "Failed to retrieve style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Style Background Color settings. Background color = $LO_COLOR_TEAL.
	_LOImpress_ShapePresStyleAreaColor($oStyle, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Style's Shadow settings.
	_LOImpress_ShapePresStyleAreaShadow($oStyle, True, $LOI_SHAPE_SHADOW_LOCATION_BOTTOM_CENTER, $LO_COLOR_GOLD, 600, 25, 35)
	If @error Then _ERROR($oDoc, "Failed to set Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Style's current settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_ShapePresStyleAreaShadow($oStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Style's Shadow settings are as follows: " & @CRLF & _
			"Is a shadow applied to Shapes with this Style? True/False: " & $avSettings[0] & @CRLF & _
			"The Shadow location is (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"The Color of the Shadow is (as a RGB Color Integer): " & $avSettings[2] & @CRLF & _
			"The distance of the shadow from the Shape is, in Hundredths of a Millimeter (HMM): " & $avSettings[3] & @CRLF & _
			"The amount of blur applied to the shadow is, in Printer's Points: " & $avSettings[4] & @CRLF & _
			"The percentage of transparency is: " & $avSettings[5])

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
