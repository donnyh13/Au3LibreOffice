#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
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

	; Modify the Image color adjust settings. Skip red adjust , Skip Green, Skip Blue, Skip Brightness, Skip Contrast, Skip Gamma, set Color mode to $LOW_COLORMODE_BLACK_WHITE
	_LOWriter_ImageColorAdjust($oImage, Null,Null,Null,Null,Null, Null, $LOW_COLORMODE_BLACK_WHITE)
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageColorAdjust($oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's color adjust settings are as follows: " & @CRLF & _
			"The image's Red adjust percentage is: " & $avSettings[0] & @CRLF & _
			"The image's Green adjust percentage is: " & $avSettings[1] & @CRLF & _
			"The image's Blue adjust percentage is: " & $avSettings[2] & @CRLF & _
			"The image's Brightness percentage is: " & $avSettings[3] & @CRLF & _
			"The image's Contrast percentage is: " & $avSettings[4] & @CRLF & _
			"The image's Gamma value is: " & $avSettings[5] & @CRLF & _
			"The image's color mode is, (See UDF constants): " & $avSettings[6] & @CRLF & _
			"Invert the image's colors? True/False: " & $avSettings[7])

	; Modify the Image color adjust settings. Skip red adjust , Skip Green, Skip Blue, Skip Brightness, Skip Contrast, Skip Gamma, set Color mode back
	; to $LOW_COLORMODE_STANDARD, set Invert to True.
	_LOWriter_ImageColorAdjust($oImage, Null,Null,Null,Null,Null, Null, $LOW_COLORMODE_STANDARD, True)
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageColorAdjust($oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's color adjust settings are as follows: " & @CRLF & _
			"The image's Red adjust percentage is: " & $avSettings[0] & @CRLF & _
			"The image's Green adjust percentage is: " & $avSettings[1] & @CRLF & _
			"The image's Blue adjust percentage is: " & $avSettings[2] & @CRLF & _
			"The image's Brightness percentage is: " & $avSettings[3] & @CRLF & _
			"The image's Contrast percentage is: " & $avSettings[4] & @CRLF & _
			"The image's Gamma value is: " & $avSettings[5] & @CRLF & _
			"The image's color mode is, (See UDF constants): " & $avSettings[6] & @CRLF & _
			"Invert the image's colors? True/False: " & $avSettings[7])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
