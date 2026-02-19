#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $bRunning

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set slide background.
	_LOImpress_SlideBackColor($oSlide, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1))
	If @error Then _ERROR($oDoc, "Failed to set Slide background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if a slideshow is currently running
	$bRunning = _LOImpress_SlideshowIsRunning($oDoc)
	If @error Then _ERROR($oDoc, "Failed to check if a Slideshow is currently running. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Is the Slideshow currently running? True/False: " & $bRunning & @CRLF & @CRLF & _
			"Press ok to start the Slideshow and check again.")

	; Begin a Slide Show
	_LOImpress_SlideshowStart($oDoc)
	If @error Then _ERROR($oDoc, "Failed to start a Slide presentation. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check again if a slideshow is currently running
	$bRunning = _LOImpress_SlideshowIsRunning($oDoc)
	If @error Then _ERROR($oDoc, "Failed to check if a Slideshow is currently running. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now is the Slideshow currently running? True/False: " & $bRunning)

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
