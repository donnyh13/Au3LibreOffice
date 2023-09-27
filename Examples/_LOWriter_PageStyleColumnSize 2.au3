
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $iMicrometers
	Local $avPageStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set Page style Column count to 4.
	_LOWriter_PageStyleColumnSettings($oPageStyle, 4)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Set Page style Column size settings for column 3, set auto width to False, and spacing for this specific column to 1/4"
	_LOWriter_PageStyleColumnSize($oPageStyle, 3, True, Null, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current settings. Return will be an array with elements in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleColumnSize($oPageStyle, 3)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Page Style's current Column size settings are as follows: " & @CRLF & _
			"Is Column width automatically adjusted? True/False: " & $avPageStyleSettings[0] & @CRLF & _
			"The Global Spacing value for the entire table, in Micrometers (If there is one): " & $avPageStyleSettings[1] & @CRLF & _
			"The Spacing value between this column and the next column to the right is, in Micrometers: " & $avPageStyleSettings[2] & @CRLF & _
			"The width of this column, in Micrometers: " & $avPageStyleSettings[3] & @CRLF & _
			"Note: This value will be different from the UI value, even when converted to Inches or Centimeters, because the returned width value is a " & _
			"relative width, not a metric width, which is why I don't know how to set this value appropriately.")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

