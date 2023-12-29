#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If @error Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	; Turn Footer on.
	_LOWriter_PageStyleFooter($oPageStyle, True)
	If @error Then _ERROR("Failed to turn Page Style footers on. Error:" & @error & " Extended:" & @extended)

	; Set Footer Border Width (all four sides) to $LOW_BORDERWIDTH_MEDIUM
	_LOWriter_PageStyleFooterBorderWidth($oPageStyle, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Set Page style Footer Border Color settings to: Top, $LOW_COLOR_ORANGE, Bottom $LOW_COLOR_BLUE, Left, $LOW_COLOR_LGRAY, Right $LOW_COLOR_BLACK
	_LOWriter_PageStyleFooterBorderColor($oPageStyle, $LOW_COLOR_ORANGE, $LOW_COLOR_BLUE, $LOW_COLOR_LGRAY, $LOW_COLOR_BLACK)
	If @error Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleFooterBorderColor($oPageStyle)
	If @error Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Footer Border Color settings are as follows: " & @CRLF & _
			"The Top Border Color is, in Long Color Format: " & $avPageStyleSettings[0] & @CRLF & _
			"The Bottom Border Color is, in Long Color Format: " & $avPageStyleSettings[1] & @CRLF & _
			"The Left Border Color is, in Long Color Format: " & $avPageStyleSettings[2] & @CRLF & _
			"The Right Border Color is, in Long Color Format: " & $avPageStyleSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
