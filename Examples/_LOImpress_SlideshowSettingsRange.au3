#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add 2 Slides
	For $i = 1 To 2
		; Insert a new slide.
		_LOImpress_SlideAdd($oDoc)
		If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Modify the play Range of the Slideshow
	_LOImpress_SlideshowSettingsRange($oDoc, $LOI_SLIDESHOW_RANGE_FROM, "Slide 2")
	If @error Then _ERROR($oDoc, "Failed to set presentation's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Slideshow's play range settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideshowSettingsRange($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Slideshow's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Slideshow's play range settings are as follows: " & @CRLF & _
			"play range is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"If present, the start slide name, or custom Slideshow name is: " & $avSettings[1])

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
