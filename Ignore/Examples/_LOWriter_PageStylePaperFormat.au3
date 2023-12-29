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

	; Set Width to $LOW_PAPER_WIDTH_10ENVELOPE, and Height to $LOW_PAPER_HEIGHT_TABLOID, Landscape to false
	_LOWriter_PageStylePaperFormat($oPageStyle, $LOW_PAPER_WIDTH_10ENVELOPE, $LOW_PAPER_HEIGHT_TABLOID, False)
	If @error Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avPageStyleSettings = _LOWriter_PageStylePaperFormat($oPageStyle)
	If @error Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Paper Format settings are as follows: " & @CRLF & _
			"The paper width, in Micrometers, is: " & $avPageStyleSettings[0] & @CRLF & _
			"The paper Height in Micrometers, is: " & $avPageStyleSettings[1] & @CRLF & _
			"Is the Page set to Landscape? True/False: " & $avPageStyleSettings[2])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
