#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set some of the Slideshow's options
	_LOImpress_SlideshowSettingsOptions($oDoc, False, True, True, True, False, False)
	If @error Then _ERROR($oDoc, "Failed to set Slideshow's option settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Slideshow's Option settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideshowSettingsOptions($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Slideshow's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Slideshow's option settings are as follows: " & @CRLF & _
			"Are slides allowed to automatically advance? True/False: " & $avSettings[0] & @CRLF & _
			"Can the slide be advanced by clicking? True/False: " & $avSettings[1] & @CRLF & _
			"Is the mouse visible? True/False: " & $avSettings[2] & @CRLF & _
			"Can the mouse be used as a pen to draw on the slide? True/False: " & $avSettings[3] & @CRLF & _
			"Are animated files (such as GIFs) played? True/False: " & $avSettings[4] & @CRLF & _
			"Is the Presentation always on top? True/False: " & $avSettings[5])

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
