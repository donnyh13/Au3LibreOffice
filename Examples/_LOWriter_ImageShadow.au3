#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $iMicrometers
	Local $avSettings
	Local $sImage = @ScriptDir & "\Extras\Plain.png"

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert an Image into the document at the Viewcursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR("Failed to insert an Image. Error:" & @error & " Extended:" & @extended)

	; Convert 1/8" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.125)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Set Image Shadow settings to: Width = 1/8", Color = $LOW_COLOR_RED, Transparent = False, Location = $LOW_SHADOW_TOP_LEFT
	_LOWriter_ImageShadow($oImage, $iMicrometers, $LOW_COLOR_RED, False, $LOW_SHADOW_TOP_LEFT)
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageShadow($oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's Shadow settings are as follows: " & @CRLF & _
			"The shadow width is, is Micrometers: " & $avSettings[0] & @CRLF & _
			"The Shdaow color is, in Long Color format: " & $avSettings[1] & @CRLF & _
			"Is the Color transparent? True/False: " & $avSettings[2] & @CRLF & _
			"The Shadow location is, (see UDF Constants): " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
