#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame, $oTextCursor, $oFrameNew

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame into the document at the ViewCursor position, and 3000x3000 Hundredths of a Millimeter (HMM) wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the frame.
	$oTextCursor = _LOWriter_FrameCreateTextCursor($oFrame)
	If @error Then _ERROR($oDoc, "Failed to create a text Cursor Object for the Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the ViewCursor into the frame.
	_LOWriter_CursorGoToRange($oViewCursor, $oTextCursor)
	If @error Then _ERROR($oDoc, "Failed to move a Cursor Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Frame Object by the ViewCursor (which is located in the Frame.)
	$oFrameNew = _LOWriter_FrameGetObjByCursor($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to Retrieve a Frame Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to delete the Frame, using the newly retrieved Object.")

	; Delete the frame using the new Frame Object.
	_LOWriter_FrameDelete($oDoc, $oFrameNew)
	If @error Then _ERROR($oDoc, "Failed to delete the Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
