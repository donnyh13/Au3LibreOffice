#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $avSettings
	Local $sImage = @ScriptDir & "\Extras\Plain.png"

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert an Image into the document at the ViewCursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert an Image. Error:" & @error & " Extended:" & @extended)

	; Modify the Image settings. Flip the image vertically, Flip the image horizontally on both left and right pages, set the angle to 55.5 degrees.
	_LOWriter_ImageModify($oImage, True, True, True, 55.5)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageModify($oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's settings are as follows: " & @CRLF & _
			"Flip the image vertically? True/False: " & $avSettings[0] & @CRLF & _
			"Flip the image horizontally on right pages? True/False: " & $avSettings[1] & @CRLF & _
			"Flip the image horizontally on left pages? True/False: " & $avSettings[2] & @CRLF & _
			"The image's angle is, in degrees: " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
