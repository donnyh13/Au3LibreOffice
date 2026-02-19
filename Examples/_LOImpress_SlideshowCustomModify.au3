#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $asCurrCustomShow
	Local $sSlides = ""
	Local $asCustomShow[3] = ["Slide 1", "Slide 2", "Slide 1"]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new slide.
	_LOImpress_SlideAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new slide.
	_LOImpress_SlideAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a custom Slideshow.
	_LOImpress_SlideshowCustomCreate($oDoc, "My_AutoIt_Show", $asCustomShow)
	If @error Then _ERROR($oDoc, "Failed to create a custom Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Slide names contained in the Custom Slideshow
	$asCurrCustomShow = _LOImpress_SlideshowCustomModify($oDoc, "My_AutoIt_Show")
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of slide names contained in the custom Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($asCurrCustomShow) - 1
		$sSlides &= $asCurrCustomShow[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There are " & UBound($asCurrCustomShow) & " Slides contained in the custom slideshow ""My_AutoIt_Show"", their names are:" & @CRLF & _
			$sSlides & @CRLF & @CRLF & _
			"Press ok to exchange ""Slide 2"" for ""Slide 3"" and append ""Slide 2"" to the end of the show.")

	ReDim $asCurrCustomShow[UBound($asCurrCustomShow) + 1]

	$asCurrCustomShow[1] = "Slide 3"
	$asCurrCustomShow[UBound($asCurrCustomShow) - 1] = "Slide 2"

	; Set the modified custom Slideshow
	_LOImpress_SlideshowCustomModify($oDoc, "My_AutoIt_Show", $asCurrCustomShow)
	If @error Then _ERROR($oDoc, "Failed to modify the custom Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Slide names contained in the Custom Slideshow
	$asCurrCustomShow = _LOImpress_SlideshowCustomModify($oDoc, "My_AutoIt_Show")
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of slide names contained in the custom Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sSlides = ""

	For $i = 0 To UBound($asCurrCustomShow) - 1
		$sSlides &= $asCurrCustomShow[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There are now " & UBound($asCurrCustomShow) & " Slides contained in the custom slideshow ""My_AutoIt_Show"", their names are:" & @CRLF & $sSlides)

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
