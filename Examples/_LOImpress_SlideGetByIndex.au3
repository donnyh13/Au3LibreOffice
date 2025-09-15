#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $iSlides

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new slide.
	_LOImpress_SlideAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new slide.
	_LOImpress_SlideAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to set the background color of all slides.")

	; Get the count of slides
	$iSlides = _LOImpress_SlidesGetCount($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve count of Slides. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To $iSlides - 1
		; Get the slide's object
		$oSlide = _LOImpress_SlideGetByIndex($oDoc, $i)
		If @error Then _ERROR($oDoc, "Failed to retrieve Slide Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Set slide background.
		_LOImpress_SlideAreaColor($oSlide, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
		If @error Then _ERROR($oDoc, "Failed to set Slide background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

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
