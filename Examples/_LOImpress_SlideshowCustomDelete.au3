#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $sShows = ""
	Local $asShows
	Local $asCustomShow[2] = ["Slide 2", "Slide 1"]
	Local $asCustomShow2[3] = ["Slide 1", "Slide 2", "Slide 1"]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new slide.
	_LOImpress_SlideAdd($oDoc)
	If @error Then _ERROR($oDoc, "Failed to Insert a new slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a custom Slideshow.
	_LOImpress_SlideshowCustomCreate($oDoc, "My_AutoIt_Show", $asCustomShow)
	If @error Then _ERROR($oDoc, "Failed to create a custom Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a custom Slideshow.
	_LOImpress_SlideshowCustomCreate($oDoc, "My_Second-AutoIt_Show", $asCustomShow2)
	If @error Then _ERROR($oDoc, "Failed to create a custom Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of custom Slideshows
	$asShows = _LOImpress_SlideshowsCustomGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of custom Slideshows. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($asShows) - 1
		$sShows &= $asShows[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There are " & UBound($asShows) & " custom slideshows contained in this document. Their names are:" & @CRLF & _
			$sShows & @CRLF & @CRLF & _
			"Press ok to delete ""My_AutoIt_Show"".")

	; Delete the first Custom Slideshow.
	_LOImpress_SlideshowCustomDelete($oDoc, $asShows[0])
	If @error Then _ERROR($oDoc, "Failed to delete a custom Slideshow. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of custom Slideshows
	$asShows = _LOImpress_SlideshowsCustomGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of custom Slideshows. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sShows = ""

	For $i = 0 To UBound($asShows) - 1
		$sShows &= $asShows[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There is " & UBound($asShows) & " custom slideshow contained in this document. Its name is:" & @CRLF & $sShows)

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
