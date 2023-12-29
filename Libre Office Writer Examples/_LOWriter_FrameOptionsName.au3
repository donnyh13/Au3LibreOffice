
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame into the document at the Viewcursor position, and 3000x3000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Insert some @CR's to move the Viewcursor down a few lines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR & @CR & @CR & @CR & @CR & @CR & @CR & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Return the cursor back to the start.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START)
	If (@error > 0) Then _ERROR("Error performing cursor Move. Error:" & @error & " Extended:" & @extended)

	;Insert another Frame into the document at the Viewcursor position, Named "AutoitTest", and 3000x3000 Micrometers wide.
	_LOWriter_FrameCreate($oDoc, $oViewCursor, "AutoitTest", 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Modify the Frame Name Option settings. Set the Lower Frame name to "AutoitTest2", Set the description to
	;"This is a Frame to demonstrate _LOWriter_FrameOptionsName.", set Previous link to the upper Frame's name,
	_LOWriter_FrameOptionsName($oDoc, $oFrame, "AutoitTest2", "This is a Frame to demonstrate _LOWriter_FrameOptionsName.", "AutoitTest")
	If (@error > 0) Then _ERROR("Failed to set Frame settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameOptionsName($oDoc, $oFrame)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame's Name Option settings are as follows: " & @CRLF & _
			"The Frame's name is: " & $avSettings[0] & @CRLF & _
			"The Frame's description is: " & $avSettings[1] & @CRLF & _
			"The previous linked frame name is (if there is one): " & $avSettings[2] & @CRLF & _
			"The next linked frame name is (if there is one): " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



