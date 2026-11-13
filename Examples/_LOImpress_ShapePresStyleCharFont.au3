#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oStyle, $oShape, $oTextCursor
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

	; Set the Style's Font to "Arial", Font size to 18, Posture (Italic) to $LOI_CHAR_POSTURE_ITALIC, and weight (Bold) to $LOI_CHAR_WEIGHT_BOLD
	_LOImpress_ShapeStyleCharFont($oStyle, "Arial", 18, $LOI_CHAR_POSTURE_ITALIC, $LOI_CHAR_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set the Style's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOImpress_ShapeStyleCharFont($oStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Style's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Style's current Font settings are as follows: " & @CRLF & _
			"Font name: " & $avSettings[0] & @CRLF & _
			"Font size: " & $avSettings[1] & @CRLF & _
			"Italic/Posture setting (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"Weight/Bold setting (See UDF Constants): " & $avSettings[3])

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
