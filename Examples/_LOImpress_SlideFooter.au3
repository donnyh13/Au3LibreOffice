#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to enable the slide's footer.")

	; Modify the Slide Footer settings, Enable Date/Time, make it variable, and make the Date and Time display in the Month, Day, Year 12Hr format, enable footer text,
	; make the footer text display "Hi from AutoIt!", and enable the current slide number field.
	_LOImpress_SlideFooter($oSlide, True, False, Null, $LOI_SLIDE_DT_FMT_MMDDYY_12H_HM_AMPM, True, "Hi from AutoIt!", True)
	If @error Then _ERROR($oDoc, "Failed to set Slide settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Slide settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideFooter($oSlide)
	If @error Then _ERROR($oDoc, "Failed to retrieve Slide settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Slide's footer settings are as follows: " & @CRLF & _
			"Is the Date/Time field enabled? True/False: " & $avSettings[0] & @CRLF & _
			"Is the Date/Time Fixed: " & $avSettings[1] & @CRLF & _
			"The custom Date/Time is (if any): " & $avSettings[2] & @CRLF & _
			"The Date/Time display format is (see UDF constants): " & $avSettings[3] & @CRLF & _
			"Is there a Footer Text enabled? True/False: " & $avSettings[4] & @CRLF & _
			"The Footer Text is (if any): " & $avSettings[5] & @CRLF & _
			"Is the Current Slide number field enabled? True/False: " & $avSettings[6])

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
