
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $iMicrometers
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame into the document at the Viewcursor position, and 3000x3000 Micrometers wide.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Set Frame Column count to 4.
	_LOWriter_FrameColumnSettings($oFrame, 4)
	If (@error > 0) Then _ERROR("Failed to modify Frame settings. Error:" & @error & " Extended:" & @extended)

	;Convert 1/4" to Micrometers
	$iMicrometers = _LOWriter_ConvertToMicrometer(.25)
	If (@error > 0) Then _ERROR("Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	;Set Frame Column size settings for column 3, set auto width to False, and spacing for this specific column to 1/4"
	_LOWriter_FrameColumnSize($oFrame, 3, True, Null, $iMicrometers)
	If (@error > 0) Then _ERROR("Failed to set Frame settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameColumnSize($oFrame, 3)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame's current Column size settings are as follows: " & @CRLF & _
			"Is Column width automatically adjusted? True/False: " & $avSettings[0] & @CRLF & _
			"The Global Spacing value for the entire frame, in Micrometers (If there is one): " & $avSettings[1] & @CRLF & _
			"The Spacing value between this column and the next column to the right is, in Micrometers: " & $avSettings[2] & @CRLF & _
			"The width of this column, in Micrometers: " & $avSettings[3] & @CRLF & _
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

