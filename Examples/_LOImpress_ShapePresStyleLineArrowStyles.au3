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

	; Convert 1/4" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(.25, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Style Arrow Style settings to: Set the Start Arrowhead to $LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, Start width = 1/4",
	; Start Center = True, Synchronize Start and End = True.
	_LOImpress_ShapePresStyleLineArrowStyles($oDoc, $oStyle, $LOI_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, $iHMM, True, True)
	If @error Then _ERROR($oDoc, "Failed to set Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_ShapePresStyleLineArrowStyles($oDoc, $oStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Style's Arrow Style settings are as follows: " & @CRLF & _
			"The Start Arrowhead Style is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Start Arrowhead Width is, in Hundredths of a Millimeter (HMM): " & $avSettings[1] & @CRLF & _
			"Is the Start Arrowhead centered on the line end? True/False: " & $avSettings[2] & @CRLF & _
			"Is the Starting and Ending Arrowhead settings synchronized? True/False: " & $avSettings[3] & @CRLF & _
			"The End Arrowhead Style is (See UDF Constants): " & $avSettings[4] & @CRLF & _
			"The End Arrowhead Width is, in Hundredths of a Millimeter (HMM): " & $avSettings[5] & @CRLF & _
			"Is the Start Arrowhead centered on the line end? True/False: " & $avSettings[6])

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
