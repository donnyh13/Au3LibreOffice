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

	; Insert some @CR's to move the ViewCursor down a few lines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & @CR & @CR & @CR & @CR & @CR & @CR & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Return the cursor back to the start.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert another Frame into the document at the ViewCursor position, Named "AutoItTest", and 3000x3000 Hundredths of a Millimeter (HMM) wide.
	_LOWriter_FrameCreate($oDoc, $oViewCursor, "AutoItTest", 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Frame Name Option settings. Set the Lower Frame name to "AutoItTest2", Set the description to
	; "This is a Frame to demonstrate _LOWriter_FrameOptionsName.", set Previous link to the upper Frame's name,
	_LOWriter_FrameOptionsName($oDoc, $oFrame, "AutoItTest2", "This is a Frame to demonstrate _LOWriter_FrameOptionsName.", "AutoItTest")
	If @error Then _ERROR($oDoc, "Failed to set Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameOptionsName($oDoc, $oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame's Name Option settings are as follows: " & @CRLF & _
			"The Frame's name is: " & $avSettings[0] & @CRLF & _
			"The Frame's description is: " & $avSettings[1] & @CRLF & _
			"The previous linked frame name is (if there is one): " & $avSettings[2] & @CRLF & _
			"The next linked frame name is (if there is one): " & $avSettings[3])

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
