#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $avSettings
	Local $sImage = @ScriptDir & "\Extras\Plain.png"

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert an Image into the document at the ViewCursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert an Image. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Image Size settings. Skip Width, set relative width to 50%, Relative to = $LOW_RELATIVE_PARAGRAPH,
	; Skip height, Relative height = 70%, Relative to = $LOW_RELATIVE_PARAGRAPH, Keep Ratio = False
	_LOWriter_ImageTypeSize($oDoc, $oImage, Null, 50, $LOW_RELATIVE_PARAGRAPH, Null, 70, $LOW_RELATIVE_PARAGRAPH, False)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageTypeSize($oDoc, $oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Image's size settings are as follows: " & @CRLF & _
			"The Image width is, in Hundredths of a Millimeter (HMM): " & $avSettings[0] & @CRLF & _
			"The Image relative width percentage is: " & $avSettings[1] & @CRLF & _
			"The width is relative to what? (See UDF Constants): " & $avSettings[2] & @CRLF & _
			"The Image height is, in Hundredths of a Millimeter (HMM): " & $avSettings[3] & @CRLF & _
			"The Image relative height percentage is: " & $avSettings[4] & @CRLF & _
			"The height is relative to what? (See UDF Constants): " & $avSettings[5] & @CRLF & _
			"Keep Height width Ratio? True/False: " & $avSettings[6])

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
