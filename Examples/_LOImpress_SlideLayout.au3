#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $iLayout

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current slide layout.
	$iLayout = _LOImpress_SlideLayout($oSlide)
	If @error Then _ERROR($oDoc, "Failed to retrieve Slide Layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Slide's current layout is (See UDF Constants): " & $iLayout & @CRLF & @CRLF & _
			"Press ok to change the slide's layout to $LOI_SLIDE_LAYOUT_TITLE_4_CONTENT.")

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_TITLE_4_CONTENT
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_TITLE_4_CONTENT)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Stop the slideshow.
	If _LOImpress_SlideShowIsRunning($oDoc) Then _LOImpress_SlideShowStop($oDoc)
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
