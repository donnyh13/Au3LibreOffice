#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $asFrames

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_ViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame into the document.
	_LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Frame names currently in the document.
	$asFrames = _LOWriter_FramesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve a list of Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If (UBound($asFrames) > 0) Then
		; Retrieve the object for the first frame listed in the Array.
		$oFrame = _LOWriter_FrameGetObjByName($oDoc, $asFrames[0])
		If @error Then _ERROR($oDoc, "Failed to retrieve a frame Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to delete the Text Frame.")

		; Delete the Frame.
		_LOWriter_FrameDelete($oDoc, $oFrame)
		If @error Then _ERROR($oDoc, "Failed to delete a frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Else
		_ERROR($oDoc, "Something went wrong, and no frames were found." & " On Line: " & @ScriptLineNumber)
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
