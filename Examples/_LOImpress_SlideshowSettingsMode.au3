#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the display mode of the Slideshow
	_LOImpress_SlideshowSettingsMode($oDoc, $LOI_SLIDESHOW_VIEW_MODE_LOOP, 3, True)
	If @error Then _ERROR($oDoc, "Failed to set presentation's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Slideshow's viewing mode settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideshowSettingsMode($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Slideshow's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Slideshow's viewing mode settings are as follows: " & @CRLF & _
			"The viewing mode is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The pause between repeating the Slideshow is (in seconds): " & $avSettings[1] & @CRLF & _
			"is the LibreOffice logo visible during the pause? True/False: " & $avSettings[2])

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
