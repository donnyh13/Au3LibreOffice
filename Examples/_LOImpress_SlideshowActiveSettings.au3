#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set slide background.
	_LOImpress_SlideBackColor($oSlide, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Slide background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Begin a Slide Show
	_LOImpress_SlideshowStart($oDoc)
	If @error Then _ERROR($oDoc, "Failed to start a Slide presentation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set some of the Active Slideshow's settings
	_LOImpress_SlideshowActiveSettings($oDoc, False, True, True, $LO_COLOR_LIME, $LOI_SLIDESHOW_PEN_WIDTH_THIN)
	If @error Then _ERROR($oDoc, "Failed to set active presentation's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Slideshow's settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideshowActiveSettings($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve presentation's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The active Slideshow's settings are as follows: " & @CRLF & _
			"Is the Presentation always on top? True/False: " & $avSettings[0] & @CRLF & _
			"Is the mouse visible? True/False: " & $avSettings[1] & @CRLF & _
			"Can the mouse be used as a pen to draw on the slide? True/False: " & $avSettings[2] & @CRLF & _
			"The color of the pen is(as a RGB Color Integer): " & $avSettings[3] & @CRLF & _
			"The width of the pen is (See UDF Constants): " & $avSettings[4])

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
