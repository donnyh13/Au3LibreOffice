#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $asImages
	Local $sImage = @ScriptDir & "\Extras\Plain.png"

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_ViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert an Image into the document at the ViewCursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert an Image. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Image names currently in the document.
	$asImages = _LOWriter_ImagesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve a list of Images. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If (UBound($asImages) > 0) Then
		; Retrieve the object for the first Image listed in the Array.
		$oImage = _LOWriter_ImageGetObjByName($oDoc, $asImages[0])
		If @error Then _ERROR($oDoc, "Failed to retrieve an Image Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to delete the Text Image.")

		; Delete the Image.
		_LOWriter_ImageDelete($oDoc, $oImage)
		If @error Then _ERROR($oDoc, "Failed to delete an Image. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Else
		_ERROR($oDoc, "Something went wrong, and no Images were found." & " On Line: " & @ScriptLineNumber)
	EndIf

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
