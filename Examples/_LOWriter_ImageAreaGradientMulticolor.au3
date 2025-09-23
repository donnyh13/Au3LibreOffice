#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oImage
	Local $sStops = "", $sImage = @ScriptDir & "\Extras\Transparent.png"
	Local $avStops

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert an Image into the document at the ViewCursor position.
	$oImage = _LOWriter_ImageInsert($oDoc, $sImage, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert an Image. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Image Gradient settings to: Preset Gradient name = $LOW_GRAD_NAME_SUNDOWN
	_LOWriter_ImageAreaGradient($oDoc, $oImage, $LOW_GRAD_NAME_SUNDOWN)
	If @error Then _ERROR($oDoc, "Failed to set Image settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Multicolor Gradient ColorStops.
	$avStops = _LOWriter_ImageAreaGradientMulticolor($oImage)
	If @error Then _ERROR($oDoc, "Failed to retrieve Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Color: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Image's Gradient ColorStops are as follows: " & @CRLF & _
			$sStops & @CRLF & @CRLF & _
			"Press ok to add a new ColorStop.")

	; Add a new ColorStop in the middle.
	_LOWriter_GradientMulticolorAdd($avStops, 3, 0.6, 1234567)
	If @error Then _ERROR($oDoc, "Failed to add a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new ColorStops.
	_LOWriter_ImageAreaGradientMulticolor($oImage, $avStops)
	If @error Then _ERROR($oDoc, "Failed to modify Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

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
