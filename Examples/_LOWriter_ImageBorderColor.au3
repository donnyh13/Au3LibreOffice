#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $sImage = @ScriptDir & "\Extras\Plain.png"
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert an Image into the document at the Viewcursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR("Failed to insert an Image. Error:" & @error & " Extended:" & @extended)

	; Set Image Border Width settings to: $LOW_BORDERWIDTH_MEDIUM on all four sides.
	_LOWriter_ImageBorderWidth($oImage, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM, $LOW_BORDERWIDTH_MEDIUM)
	If @error Then _ERROR("Failed to modify Image settings. Error:" & @error & " Extended:" & @extended)

	; Modify the Image Border Color settings to: Top, $LOW_COLOR_ORANGE, Bottom $LOW_COLOR_BLUE, Left, $LOW_COLOR_MAGENTA, Right $LOW_COLOR_BLACK
	_LOWriter_ImageBorderColor($oImage, $LOW_COLOR_ORANGE, $LOW_COLOR_BLUE, $LOW_COLOR_MAGENTA, $LOW_COLOR_BLACK)
	If @error Then _ERROR("Failed to set Image settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Image settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_ImageBorderColor($oImage)
	If @error Then _ERROR("Failed to retrieve Image settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Image's Border Color settings are as follows: " & @CRLF & _
			"The Top Border Color is, in Long Color Format: " & $avSettings[0] & @CRLF & _
			"The Bottom Border Color is, in Long Color Format: " & $avSettings[1] & @CRLF & _
			"The Left Border Color is, in Long Color Format: " & $avSettings[2] & @CRLF & _
			"The Right Border Color is, in Long Color Format: " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
