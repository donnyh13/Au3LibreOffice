#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $avSettings
	Local $iMicrometers, $iMicrometers2
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

	; Modify the Image size. Set Scale Width to 65%, Scale Height to 45%
	_LOWriter_ImageSize($oImage, 65, 45)
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageSize($oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's sizes are as follows: " & @CRLF & _
			"The Image's scale width percentage is: " & $avSettings[0] & @CRLF & _
			"The Image's scale height percentage is: " & $avSettings[1] & @CRLF & _
			"The Image's width, in Micrometers, is: " & $avSettings[2] & @CRLF & _
			"The Image's height, in Micrometers, is: " & $avSettings[3] & @CRLF & _
			"Is the Image currently at its original size? True/False: " & $avSettings[4])

	MsgBox($MB_OK,"", "Press ok to return the image to its original size.")

	; Return the Image to its original size.
	_LOWriter_ImageSize($oImage, Null, Null,Null, Null, True)
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageSize($oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's sizes are as follows: " & @CRLF & _
			"The Image's scale width percentage is: " & $avSettings[0] & @CRLF & _
			"The Image's scale height percentage is: " & $avSettings[1] & @CRLF & _
			"The Image's width, in Micrometers, is: " & $avSettings[2] & @CRLF & _
			"The Image's height, in Micrometers, is: " & $avSettings[3] & @CRLF & _
			"Is the Image currently at its original size? True/False: " & $avSettings[4])

	MsgBox($MB_OK,"", "Press ok size the image again.")

	; Convert 4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(4)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Convert 7" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(7)
	If @error Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Set the Image's size. Skip Scale Width, Skip Scale Height, Set Width to 4", height to 7"
	_LOWriter_ImageSize($oImage, Null, Null,$iMicrometers, $iMicrometers2)
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageSize($oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's sizes are as follows: " & @CRLF & _
			"The Image's scale width percentage is: " & $avSettings[0] & @CRLF & _
			"The Image's scale height percentage is: " & $avSettings[1] & @CRLF & _
			"The Image's width, in Micrometers, is: " & $avSettings[2] & @CRLF & _
			"The Image's height, in Micrometers, is: " & $avSettings[3] & @CRLF & _
			"Is the Image currently at its original size? True/False: " & $avSettings[4])

	MsgBox($MB_OK,"", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
