#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTextCursor
	Local $avShapes

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set slide background.
	_LOImpress_SlideBackColor($oSlide, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Slide background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of Textboxes in the slide.
	$avShapes = _LOImpress_SlideShapesGetList($oSlide, BitOR($LOI_SHAPE_TYPE_TEXTBOX, $LOI_SHAPE_TYPE_TEXTBOX_TITLE, $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE))
	If @error Or (@extended = 0) Then _ERROR($oDoc, "Failed to retrieve Shapes, or no Shapes present in Slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the Textbox.
	$oTextCursor = _LOImpress_DrawShapeTextboxCreateTextCursor($avShapes[0][0])
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert the Slide's number.
	_LOImpress_CursorInsertString($oTextCursor, "Slide #1")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add 3 Slides
	For $i = 1 To 3
		; Insert a new slide.
		$oSlide = _LOImpress_SlideAdd($oDoc)
		If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set slide background.
		_LOImpress_SlideBackColor($oSlide, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
		If @error Then _ERROR($oDoc, "Failed to set Slide background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Change the Slide's layout to $LOI_SLIDE_LAYOUT_TITLE_ONLY
		_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_TITLE_ONLY)
		If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Retrieve an Array of Textboxes in the slide.
		$avShapes = _LOImpress_SlideShapesGetList($oSlide, BitOR($LOI_SHAPE_TYPE_TEXTBOX, $LOI_SHAPE_TYPE_TEXTBOX_TITLE, $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE))
		If @error Or (@extended = 0) Then _ERROR($oDoc, "Failed to retrieve Shapes, or no Shapes present in Slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Create a Text Cursor in the Textbox.
		$oTextCursor = _LOImpress_DrawShapeTextboxCreateTextCursor($avShapes[0][0])
		If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Insert the Slide's number.
		_LOImpress_CursorInsertString($oTextCursor, "Slide #" & $i + 1)
		If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Start the Slideshow
	_LOImpress_SlideshowStart($oDoc)
	If @error Then _ERROR($oDoc, "Failed to start the Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have started the Slideshow from the beginning." & @CRLF & @CRLF & _
			"Press ok to stop the Slideshow, and start it from ""Slide 3"".")

	; Stop the Slideshow.
	_LOImpress_SlideshowStop($oDoc)
	If @error Then _ERROR($oDoc, "Failed to stop the Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Start the Slideshow from "Slide 3"
	_LOImpress_SlideshowStart($oDoc, False, "Slide 3")
	If @error Then _ERROR($oDoc, "Failed to start the Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
