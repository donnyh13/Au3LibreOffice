#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame into the document at the ViewCursor position, and 3000x3000 Hundredths of a Millimeter (HMM) wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame wrap option settings. First Paragraph = True, Skip InBackground, Allow overlap = False
	_LOWriter_FrameWrapOptions($oFrame, True, Null, False)
	If @error Then _ERROR($oDoc, "Failed to set Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameWrapOptions($oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame's wrap option settings are as follows: " & @CRLF & _
			"Create a new paragraph below the frame? True/False: " & $avSettings[0] & @CRLF & _
			"Place the frame in the background? True/False: " & $avSettings[1] & @CRLF & _
			"Allow multiple frames to overlap? True/False: " & $avSettings[2])

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
