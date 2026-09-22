#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $sMasterSlides = ""
	Local $asMasterSlides

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add 4 new Master Slides
	For $i = 1 To 5
		_LOImpress_SlideMasterAdd($oDoc, Null, "Au3 Master " & $i)
		If @error Then _ERROR($oDoc, "Failed to Insert a new Master slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

	; Retrieve an Array of all Master slides in the Document.
	$asMasterSlides = _LOImpress_SlideMastersGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Master Slide names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Cycle through and list all the master slides.
	For $i = 0 To @extended - 1
		$sMasterSlides &= $asMasterSlides[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "This Document contains the following Master Slides:" & @CRLF & $sMasterSlides)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
