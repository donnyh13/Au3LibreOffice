#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Frame into the document at the Viewcursor position, and 3000x3000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	; Modify the Frame Size settings. Skip Width, set relative width to 50%, Relative to = $LOW_RELATIVE_PARAGRAPH,
	; AutoWidth = False, Skip height, Relative height = 70%, Relative to = $LOW_RELATIVE_PARAGRAPH, AutoHeight = False, Keep Ratio = False
	_LOWriter_FrameTypeSize($oDoc, $oFrame, Null, 50, $LOW_RELATIVE_PARAGRAPH, Null, Null, 70, $LOW_RELATIVE_PARAGRAPH, False, False)
	If @error Then _ERROR("Failed to set Frame settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameTypeSize($oDoc, $oFrame)
	If @error Then _ERROR("Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame's size settings are as follows: " & @CRLF & _
			"The Frame's width is, in Micrometers: " & $avSettings[0] & @CRLF & _
			"The frame's relative width percentage is: " & $avSettings[1] & @CRLF & _
			"The width is relative to what? (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"Automatic width? True/False: " & $avSettings[3] & @CRLF & _
			"The Frame's height is, in Micrometers: " & $avSettings[4] & @CRLF & _
			"The frame's relative height percentage is: " & $avSettings[5] & @CRLF & _
			"The height is relative to what? (See UDF Constants): " & $avSettings[6] & @CRLF & _
			"Automatic Height? True/False: " & $avSettings[7] & @CRLF & _
			"Keep Height width Ratio? True/False: " & $avSettings[8])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
