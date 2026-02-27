#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $avSettings
	Local $asSounds

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all built-in sounds available.
	$asSounds = _LOImpress_SlideSoundsGetNames()
	If @error Then _ERROR($oDoc, "Failed to retrieve Slide names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If (UBound($asSounds) = 0) Then ; If not sounds found, add an empty entry to the example doesn't fail.
		ReDim $asSounds[1]
		$asSounds[0] = ""
	EndIf

	; Set some of the Slide's Transition settings
	_LOImpress_SlideTransition($oSlide, $LOI_SLIDE_TRANSITION_COVER_TOP_LEFT_TO_BOTTOM_RIGHT, 3.5, $asSounds[0], False, 7)
	If @error Then _ERROR($oDoc, "Failed to set Slide's Transition settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Slide's transition settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideTransition($oSlide)
	If @error Then _ERROR($oDoc, "Failed to retrieve slide's Transition settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Slide's Transition settings are as follows: " & @CRLF & _
			"The Transition effect is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The duration of the effect is, in Seconds: " & $avSettings[1] & @CRLF & _
			"The sound file played during the effect (if any) is: " & $avSettings[2] & @CRLF & _
			"Is the sound looped? True/False: " & $avSettings[3] & @CRLF & _
			"The time in seconds before the slide automatically advances is: " & $avSettings[4])

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
