#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $iHMM, $iHMM2
	Local $avSettings
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

	; Convert 1" to Hundredths of a Millimeter (HMM)
	$iHMM = _LO_UnitConvert(1, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2" to Hundredths of a Millimeter (HMM)
	$iHMM2 = _LO_UnitConvert(.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Image crop settings. Set left and Right to 1", Top and Bottom to 1/2". Keep scale = false, keep image size.
	_LOWriter_ImageCrop($oImage, $iHMM, $iHMM, $iHMM2, $iHMM2, False)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageCrop($oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Image's crop settings are as follows: " & @CRLF & _
			"The image's left crop value, in Hundredths of a Millimeter (HMM) is: " & $avSettings[0] & @CRLF & _
			"The image's Right crop value, in Hundredths of a Millimeter (HMM) is: " & $avSettings[1] & @CRLF & _
			"The image's Top crop value, in Hundredths of a Millimeter (HMM) is: " & $avSettings[2] & @CRLF & _
			"The image's Bottom crop value, in Hundredths of a Millimeter (HMM) is: " & $avSettings[3] & @CRLF & _
			"Keep the image's Scale? True/False: " & $avSettings[4])

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
