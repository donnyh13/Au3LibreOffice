#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTextCursor
	Local $avShapes
	Local $iCount, $iCurrSlideIndex
	Local $bPaused

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
	_LOImpress_CursorInsertString($oTextCursor, "Slide #1")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add 4 Slides
	For $i = 1 To 4
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

	; Begin a Slide Show
	_LOImpress_SlideshowStart($oDoc)
	If @error Then _ERROR($oDoc, "Failed to start a Slide presentation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve a count of Slides in the Slideshow.
	$iCount = _LOImpress_SlideshowPresentationControl($oDoc, $LOI_SLIDESHOW_PRES_QUERY_GET_SLIDE_COUNT)
	If @error Then _ERROR($oDoc, "Failed to query Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Advance one slide
	_LOImpress_SlideshowPresentationControl($oDoc, $LOI_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_SLIDE)
	If @error Then _ERROR($oDoc, "Failed to command Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Currently shown Slide's Index
	$iCurrSlideIndex = _LOImpress_SlideshowPresentationControl($oDoc, $LOI_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE_INDEX)
	If @error Then _ERROR($oDoc, "Failed to query Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There are " & $iCount & " Slides in this presentation. The Slideshow is currently on slide index number (0 based): " & $iCurrSlideIndex & @CRLF & @CRLF & _
			"Press ok to go to a blank pause screen.")

	; Go to a blank pause screen
	_LOImpress_SlideshowPresentationControl($oDoc, $LOI_SLIDESHOW_PRES_COMMAND_ACTIVATE_BLANK_SCREEN, $LO_COLOR_MAGENTA)
	If @error Then _ERROR($oDoc, "Failed to command Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; See if the Slideshow is paused
	$bPaused = _LOImpress_SlideshowPresentationControl($oDoc, $LOI_SLIDESHOW_PRES_QUERY_IS_PAUSED)
	If @error Then _ERROR($oDoc, "Failed to query Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the Slideshow currently paused? True/False: " & $bPaused)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Stop the slideshow.
	If _LOImpress_SlideshowIsRunning($oDoc) Then _LOImpress_SlideshowStop($oDoc)
	If @error Then _ERROR($oDoc, "Failed to stop the active slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
