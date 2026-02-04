#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $avSettings
	Local $iHMM, $iHMM2
	Local $sImage = @ScriptDir & "\Extras\Plain.png"

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert an Image into the document at the ViewCursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert an Image. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Image size. Set Scale Width to 65%, Scale Height to 45%
	_LOWriter_ImageSize($oImage, 65, 45)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageSize($oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Image's sizes are as follows: " & @CRLF & _
			"The Image's scale width percentage is: " & $avSettings[0] & @CRLF & _
			"The Image's scale height percentage is: " & $avSettings[1] & @CRLF & _
			"The Image's width, in Hundredths of a Millimeter (HMM), is: " & $avSettings[2] & @CRLF & _
			"The Image's height, in Hundredths of a Millimeter (HMM), is: " & $avSettings[3] & @CRLF & _
			"Is the Image currently at its original size? True/False: " & $avSettings[4])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to return the image to its original size.")

	; Return the Image to its original size.
	_LOWriter_ImageSize($oImage, Null, Null, Null, Null, True)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageSize($oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Image's sizes are as follows: " & @CRLF & _
			"The Image's scale width percentage is: " & $avSettings[0] & @CRLF & _
			"The Image's scale height percentage is: " & $avSettings[1] & @CRLF & _
			"The Image's width, in Hundredths of a Millimeter (HMM), is: " & $avSettings[2] & @CRLF & _
			"The Image's height, in Hundredths of a Millimeter (HMM), is: " & $avSettings[3] & @CRLF & _
			"Is the Image currently at its original size? True/False: " & $avSettings[4])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok size the image again.")

	; Convert 4" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(4, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 7" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(7, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Image's size. Skip Scale Width, Skip Scale Height, Set Width to 4", height to 7"
	_LOWriter_ImageSize($oImage, Null, Null, $iHMM, $iHMM2)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageSize($oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Image's sizes are as follows: " & @CRLF & _
			"The Image's scale width percentage is: " & $avSettings[0] & @CRLF & _
			"The Image's scale height percentage is: " & $avSettings[1] & @CRLF & _
			"The Image's width, in Hundredths of a Millimeter (HMM), is: " & $avSettings[2] & @CRLF & _
			"The Image's height, in Hundredths of a Millimeter (HMM), is: " & $avSettings[3] & @CRLF & _
			"Is the Image currently at its original size? True/False: " & $avSettings[4])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
