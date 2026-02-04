#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iFillStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page Style Fill Style
	$iFillStyle = _LOWriter_PageStyleAreaFillStyle($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

	; Set Page style Background color to $LO_COLOR_LIME.
	_LOWriter_PageStyleAreaColor($oPageStyle, $LO_COLOR_LIME)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page Style Fill Style
	$iFillStyle = _LOWriter_PageStyleAreaFillStyle($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Fill Style is: " & $iFillStyle & @CRLF & _
			"The result will be one of the following Constants:" & @CRLF & _
			"$LOW_AREA_FILL_STYLE_OFF, 0 Fill Style is off/ no background is applied." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_SOLID, 1 Fill Style is a solid color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_GRADIENT, 2 Fill Style is a gradient color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_HATCH, 3 Fill Style is a Hatch style color." & @CRLF & _
			"$LOW_AREA_FILL_STYLE_BITMAP, 4 Fill Style is a Bitmap.")

	; Set Page style Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_TEAL_TO_BLUE
	_LOWriter_PageStyleAreaGradient($oDoc, $oPageStyle, $LOW_GRAD_NAME_TEAL_TO_BLUE)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page Style Fill Style
	$iFillStyle = _LOWriter_PageStyleAreaFillStyle($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Fill Style is: " & $iFillStyle & @CRLF & _
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
