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

	; Modify the Frame Size settings. Skip Width, set relative width to 50%, Relative to = $LOW_RELATIVE_PARAGRAPH,
	; AutoWidth = False, Skip height, Relative height = 70%, Relative to = $LOW_RELATIVE_PARAGRAPH, AutoHeight = False, Keep Ratio = False
	_LOWriter_FrameTypeSize($oDoc, $oFrame, Null, 50, $LOW_RELATIVE_PARAGRAPH, Null, Null, 70, $LOW_RELATIVE_PARAGRAPH, False, False)
	If @error Then _ERROR($oDoc, "Failed to set Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameTypeSize($oDoc, $oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame's size settings are as follows: " & @CRLF & _
			"The Frame's width is, in Hundredths of a Millimeter (HMM): " & $avSettings[0] & @CRLF & _
			"The frame's relative width percentage is: " & $avSettings[1] & @CRLF & _
			"The width is relative to what? (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"Automatic width? True/False: " & $avSettings[3] & @CRLF & _
			"The Frame's height is, in Hundredths of a Millimeter (HMM): " & $avSettings[4] & @CRLF & _
			"The frame's relative height percentage is: " & $avSettings[5] & @CRLF & _
			"The height is relative to what? (See UDF Constants): " & $avSettings[6] & @CRLF & _
			"Automatic Height? True/False: " & $avSettings[7] & @CRLF & _
			"Keep Height width Ratio? True/False: " & $avSettings[8])

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
