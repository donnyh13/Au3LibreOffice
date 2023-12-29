#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asImages
	Local $sImage = @ScriptDir & "\Extras\Plain.png"

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert an Image into the document at the Viewcursor position.
	_LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR("Failed to insert an Image. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of Image names currently in the document.
	$asImages = _LOWriter_ImagesGetNames($oDoc)
	If @error Then _ERROR("Failed to retrieve a list of Images. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The document's Image names are:" & @CR & @CR)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asImages) -1
		; Insert the Image names.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asImages[$i] & @CR)
		If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
