#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set slide background.
	_LOImpress_SlideBackColor($oSlide, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Slide background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add 3 Slides
	For $i = 1 To 3
		; Insert a new slide.
		$oSlide = _LOImpress_SlideAdd($oDoc)
		If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set slide background.
		_LOImpress_SlideBackColor($oSlide, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
		If @error Then _ERROR($oDoc, "Failed to set Slide background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Start the Slideshow
	_LOImpress_SlideShowStart($oDoc)
	If @error Then _ERROR($oDoc, "Failed to start the Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have started the Slideshow from the beginning. Press ok to stop the Slideshow.")

	; Stop the Slideshow.
	_LOImpress_SlideShowStop($oDoc)
	If @error Then _ERROR($oDoc, "Failed to stop the Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
