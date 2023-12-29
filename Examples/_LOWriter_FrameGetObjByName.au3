
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $asFrames

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame into the document.
	_LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Retrieve an array of Frame names currently in the document.
	$asFrames = _LOWriter_FramesGetNames($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve a list of Frame. Error:" & @error & " Extended:" & @extended)

	If (UBound($asFrames) > 0) Then

		;Retrieve the object for the first frame listed in the Array.
		$oFrame = _LOWriter_FrameGetObjByName($oDoc, $asFrames[0])
		If (@error > 0) Then _ERROR("Failed to retrieve a frame Object. Error:" & @error & " Extended:" & @extended)

		MsgBox($MB_OK, "", "Press ok to delete the Text Frame.")

		;Delete the Frame.
		_LOWriter_FrameDelete($oDoc, $oFrame)
		If (@error > 0) Then _ERROR("Failed to delete a frame. Error:" & @error & " Extended:" & @extended)

	Else
		_ERROR("Something went wrong, and no frames were found.")
	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

