
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers, $iMicrometers2, $iMicrometers3
	Local $avPageStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	;Convert 1" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(1)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Convert 1/2" to Micrometers
	$iMicrometers2 = _LOWriter_ConvertToMicrometer(.5)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers3 = _LOWriter_ConvertToMicrometer(.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;If Libre Office version is higher or equal to 7.2 then set Gutter margin.
	If (_LOWriter_VersionGet(True) >= 7.2) Then
		;Set Left and Right margins to 1", Top and Bottom Margins to 1/2" and Gutter Margin to 1/4".
		_LOWriter_PageStyleMargins($oPageStyle, $iMicrometers, $iMicrometers, $iMicrometers2, $iMicrometers2, $iMicrometers3)
		If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	Else ;Set all other margins, except the Gutter margin.
		;Set Left and Right margins to 1", Top and Bottom Margins to 1/2".
		_LOWriter_PageStyleMargins($oPageStyle, $iMicrometers, $iMicrometers, $iMicrometers2, $iMicrometers2)
		If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)
	EndIf

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avPageStyleSettings = _LOWriter_PageStyleMargins($oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	;If Libre Office version is higher or equal to 7.2 then display the Gutter margin setting.
	If (_LOWriter_VersionGet(True) >= 7.2) Then
		MsgBox($MB_OK, "", "The Page Style's current Margin settings are as follows: " & @CRLF & _
				"The Left page margin, in Micrometers, is: " & $avPageStyleSettings[0] & @CRLF & _
				"The Right page margin, in Micrometers, is: " & $avPageStyleSettings[1] & @CRLF & _
				"The Top page margin, in Micrometers, is: " & $avPageStyleSettings[2] & @CRLF & _
				"The Bottom page margin, in Micrometers, is: " & $avPageStyleSettings[3] & @CRLF & _
				"The Gutter page margin, in Micrometers, is: " & $avPageStyleSettings[4])

	Else ; Display all other margin settings, except the Gutter margin.
		MsgBox($MB_OK, "", "The Page Style's current Margin settings are as follows: " & @CRLF & _
				"The Left page margin, in Micrometers, is: " & $avPageStyleSettings[0] & @CRLF & _
				"The Right page margin, in Micrometers, is: " & $avPageStyleSettings[1] & @CRLF & _
				"The Top page margin, in Micrometers, is: " & $avPageStyleSettings[2] & @CRLF & _
				"The Bottom page margin, in Micrometers, is: " & $avPageStyleSettings[3])
	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

