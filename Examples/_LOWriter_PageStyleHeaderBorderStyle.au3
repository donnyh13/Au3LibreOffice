#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	; Turn Header on.
	_LOWriter_PageStyleHeader($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn Page Style headers on. Error:" & @error & " Extended:" & @extended)

	; Set Header Border Width (all four sides) to $LOW_BORDERWIDTH_MEDIUM
	_LOWriter_PageStyleHeaderBorderWidth($oPageStyle, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Set Page style Header Border Style settings to: Top = $LOW_BORDERSTYLE_DASH_DOT_DOT, Bottom = $LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP
	; Left = $LOW_BORDERSTYLE_DOUBLE, Right = $LOW_BORDERSTYLE_DASHED
	_LOWriter_PageStyleHeaderBorderStyle($oPageStyle, $LOW_BORDERSTYLE_DASH_DOT_DOT, $LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP, $LOW_BORDERSTYLE_DOUBLE, $LOW_BORDERSTYLE_DASHED)
	If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleHeaderBorderStyle($oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Header Border Style settings are as follows: " & @CRLF & _
			"The Top Border Style is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
			"The Bottom Border Style is, (see UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
			"The Left Border Style is, (see UDF constants): " & $avPageStyleSettings[2] & @CRLF & _
			"The Right Border Style is, (see UDF constants): " & $avPageStyleSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
