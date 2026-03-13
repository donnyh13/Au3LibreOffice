#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $sImage = @ScriptDir & "\Extras\Plain.png"
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert an Image into the document at the ViewCursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert an Image. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set Image Border Width settings to: $LOW_BORDER_WIDTH_MEDIUM on all four sides.
	_LOWriter_ImageBorderWidth($oImage, $LOW_BORDER_WIDTH_MEDIUM, $LOW_BORDER_WIDTH_MEDIUM, $LOW_BORDER_WIDTH_MEDIUM, $LOW_BORDER_WIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to modify Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Image Border Style settings to: Top = $LOW_BORDER_STYLE_DASH_DOT_DOT, Bottom = $LOW_BORDER_STYLE_THICKTHIN_MEDIUMGAP
	; Left = $LOW_BORDER_STYLE_DOUBLE, Right = $LOW_BORDER_STYLE_DASHED
	_LOWriter_ImageBorderStyle($oImage, $LOW_BORDER_STYLE_DASH_DOT_DOT, $LOW_BORDER_STYLE_THICKTHIN_MEDIUMGAP, $LOW_BORDER_STYLE_DOUBLE, $LOW_BORDER_STYLE_DASHED)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageBorderStyle($oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Image's Border Style settings are as follows: " & @CRLF & _
			"The Top Border Style is, (see UDF constants): " & $avSettings[0] & @CRLF & _
			"The Bottom Border Style is, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The Left Border Style is, (see UDF constants): " & $avSettings[2] & @CRLF & _
			"The Right Border Style is, (see UDF constants): " & $avSettings[3])

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
