
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers, $iMicrometers2
	Local $avPageStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	;Set Page style Border Width settings to: $LOW_BORDERWIDTH_MEDIUM on all four sides.
	_LOWriter_PageStyleBorderWidth($oPageStyle, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set Page style Header settings to: Header on = True, Same content on left and right pages = False, Same content on the first page = Ture,
	;Left & Right margins = 1/4", Spacing between Header content and Page content = 1/2", Dynamic spacing = False, Skip Height and set AutoHeight to True.
	_LOWriter_PageStyleHeader($oPageStyle, True, False, True, $iMicrometers, $iMicrometers, $iMicrometers2, False, Null, True)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleHeader($oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Header settings are as follows: " & @CRLF & _
			"Is the Header on for this PageStyle? True/False: " & $avPageStyleSettings[0] & @CRLF & _
			"Is the content on Left and Right pages the same? True/False: " & $avPageStyleSettings[1] & @CRLF & _
			"Is the content on the first page the same? True/False: " & $avPageStyleSettings[2] & @CRLF & _
			"The Left Margin Width is, in Micrometers: " & $avPageStyleSettings[3] & @CRLF & _
			"The Right Margin Width is, in Micrometers: " & $avPageStyleSettings[4] & @CRLF & _
			"The Spacing between the Header contents and the Page contents, in Micrometers: " & $avPageStyleSettings[5] & @CRLF & _
			"Is the Spacing between the Header contents and the Page contents automatically adjusted? True/False: " & $avPageStyleSettings[6] & @CRLF & _
			"The height of the Header, in Micrometers: " & $avPageStyleSettings[7] & @CRLF & _
			"IS the height of the Header automatically adjusted? True/False: " & $avPageStyleSettings[8])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

