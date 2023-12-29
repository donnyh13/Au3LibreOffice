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

	; Modify the Image Name Option settings. Set the Image name to "AutoitTest", Set the Alternate text to "This is a Plain Image", Set the description to
	; "This is an Image to demonstrate _LOWriter_ImageOptionsName."
	_LOWriter_ImageOptionsName($oDoc, $oImage, "AutoitTest", "This is a Plain Image", "This is an Image to demonstrate _LOWriter_ImageOptionsName.")
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageOptionsName($oDoc, $oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's Name Option settings are as follows: " & @CRLF & _
			"The Image's name is: " & $avSettings[0] & @CRLF & _
			"The Image's Alternate Text is: " & $avSettings[1] & @CRLF & _
			"The Image's description is: " & $avSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
