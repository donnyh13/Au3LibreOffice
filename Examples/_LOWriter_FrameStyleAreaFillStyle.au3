#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame, $oFrameStyle
	Local $iFillStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Frame Style named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame into the document for demonstration.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to create a Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Frame's style to my created style, "Test Style"
	_LOWriter_FrameStyleCurrent($oDoc, $oFrame, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to set Frame style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Frame Style's current Fill Style
	$iFillStyle = _LOWriter_FrameStyleAreaFillStyle($oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Fill Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame Style's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

	; Modify the Frame Style Background Color settings. Background color = $LO_COLOR_TEAL
	_LOWriter_FrameStyleAreaColor($oFrameStyle, $LO_COLOR_TEAL)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Frame Style's current Fill Style
	$iFillStyle = _LOWriter_FrameStyleAreaFillStyle($oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Fill Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame Style's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

	; Modify the Frame Style Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_TEAL_TO_BLUE
	_LOWriter_FrameStyleAreaGradient($oDoc, $oFrameStyle, $LOW_GRAD_NAME_TEAL_TO_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Frame Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Frame Style's current Fill Style
	$iFillStyle = _LOWriter_FrameStyleAreaFillStyle($oFrame)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame Fill Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Frame Style's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

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
