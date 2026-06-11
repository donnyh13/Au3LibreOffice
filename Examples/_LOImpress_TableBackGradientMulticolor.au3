#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTable
	Local $sStops = ""
	Local $avStops

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active slide.
	$oSlide = _LOImpress_SlideCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active slide. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Change the Slide's layout to $LOI_SLIDE_LAYOUT_BLANK
	_LOImpress_SlideLayout($oSlide, $LOI_SLIDE_LAYOUT_BLANK)
	If @error Then _ERROR($oDoc, "Failed to modify Slide layout. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Table.
	$oTable = _LOImpress_TableInsert($oSlide, 5000, 4000, 3, 3)
	If @error Then _ERROR($oDoc, "Failed to insert a Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Table Gradient settings to: Preset Gradient name = $LOI_GRAD_NAME_SUNDOWN
	_LOImpress_TableBackGradient($oTable, $LOI_GRAD_NAME_SUNDOWN)
	If @error Then _ERROR($oDoc, "Failed to set Table settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an array of Multicolor Gradient ColorStops.
	$avStops = _LOImpress_TableBackGradientMulticolor($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($avStops) - 1
		$sStops &= "ColorStop offset: " & $avStops[$i][0] & " | " & @TAB & "ColorStop Color: " & $avStops[$i][1] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Table's Gradient ColorStops are as follows: " & @CRLF & _
			$sStops & @CRLF & @CRLF & _
			"Press ok to add a new ColorStop.")

	; Add a new ColorStop in the middle.
	_LO_GradientMulticolorAdd($avStops, 3, 0.6, 1234567)
	If @error Then _ERROR($oDoc, "Failed to add a ColorStop. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the new ColorStops.
	_LOImpress_TableBackGradientMulticolor($oTable, $avStops)
	If @error Then _ERROR($oDoc, "Failed to modify Multicolor Gradient settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc
