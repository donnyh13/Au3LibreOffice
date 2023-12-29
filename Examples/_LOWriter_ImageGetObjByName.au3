#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $asImages
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

	; Retrieve an array of Image names currently in the document.
	$asImages = _LOWriter_ImagesGetNames($oDoc)
	If @error Then _ERROR("Failed to retrieve a list of Images. Error:" & @error & " Extended:" & @extended)

	If (UBound($asImages) > 0) Then

		; Retrieve the object for the first Image listed in the Array.
		$oImage = _LOWriter_ImageGetObjByName($oDoc, $asImages[0])
		If @error Then _ERROR("Failed to retrieve an Image Object. Error:" & @error & " Extended:" & @extended)

		MsgBox($MB_OK, "", "Press ok to delete the Text Image.")

		; Delete the Image.
		_LOWriter_ImageDelete($oDoc, $oImage)
		If @error Then _ERROR("Failed to delete an Image. Error:" & @error & " Extended:" & @extended)

	Else
		_ERROR("Something went wrong, and no Images were found.")
	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
