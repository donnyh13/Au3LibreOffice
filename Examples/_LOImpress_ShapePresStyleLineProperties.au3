#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oStyle, $oShape, $oTextCursor
	Local $iHMM
	Local $avShapes, $avSettings

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

	; Retrieve an Array of Textboxes in the current slide.
	$avShapes = _LOImpress_ShapesGetList($oSlide, BitOR($LOI_SHAPE_TYPE_TEXTBOX, $LOI_SHAPE_TYPE_TEXTBOX_TITLE, $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE))
	If @error Then _ERROR($oDoc, "Failed to retrieve Shapes in slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Get the Object for the first Textbox found.
	$oShape = $avShapes[0][0]

	; Create a Text Cursor in the Textbox.
	$oTextCursor = _LOImpress_ShapeCreateTextCursor($oShape)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOImpress_CursorInsertString($oTextCursor, "Created by AutoIt!")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/8" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(.125, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Style Line Properties settings to: Set the Line Style to $LOI_SHAPE_LINE_STYLE_3_DASHES_3_DOTS, Line Color to $LO_COLOR_MAGENTA,
	; Width = 1/8", Transparency = 50%, Corner Style = $LOI_SHAPE_LINE_JOINT_BEVEL, Cap Style = $LOI_SHAPE_LINE_CAP_FLAT
	_LOImpress_ShapePresStyleLineProperties($oDoc, $oStyle, $LOI_SHAPE_LINE_STYLE_3_DASHES_3_DOTS, $LO_COLOR_MAGENTA, $iHMM, 50, $LOI_SHAPE_LINE_JOINT_BEVEL, $LOI_SHAPE_LINE_CAP_FLAT)
	If @error Then _ERROR($oDoc, "Failed to set Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_ShapePresStyleLineProperties($oDoc, $oStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Style's Line Properties settings are as follows: " & @CRLF & _
			"The Line Style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Line color is (as a RGB Color Integer): " & $avSettings[1] & @CRLF & _
			"The Line's Width is, in Hundredths of a Millimeter (HMM): " & $avSettings[2] & @CRLF & _
			"The Line's transparency percentage is: " & $avSettings[3] & @CRLF & _
			"The Line Corner Style is, (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The Line Cap Style is, (See UDF Constants): " & $avSettings[5])

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
