
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

	;Modify the Frame Option settings. Set Protect content to True, Protect Position to True, Protect size to True, Vertical alignment to
	;$LOW_TXT_ADJ_VERT_CENTER, Edit in Read-Only to True, Print to False, Text direction to $LOW_TXT_DIR_TB_LR
	_LOWriter_FrameOptions($oFrame, True, True, True, $LOW_TXT_ADJ_VERT_CENTER, True, False, $LOW_TXT_DIR_TB_LR)
	If (@error > 0) Then _ERROR("Failed to set Frame settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameOptions($oFrame)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame's option settings are as follows: " & @CRLF & _
			"Protect the Frame's contents from changes? True/False: " & $avSettings[0] & @CRLF & _
			"Protect the Frame's position from changes? True/False: " & $avSettings[1] & @CRLF & _
			"Protect the Frame's Size from changes? True/False: " & $avSettings[2] & @CRLF & _
			"The Vertical alingment of the frame is, (see UDF constants): " & $avSettings[3] & @CRLF & _
			"Allow the Frame's contents to be changed in Read-Only mode? True/False: " & $avSettings[4] & @CRLF & _
			"Print this frame when the document is printed? True/False: " & $avSettings[5] & @CRLF & _
			"The text direction for this frame is, (See UDF constants): " & $avSettings[6])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



