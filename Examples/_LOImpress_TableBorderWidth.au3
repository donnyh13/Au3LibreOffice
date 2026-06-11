#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTable
	Local $avSettings

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

	; Set the Border width.
	_LOImpress_TableBorderWidth($oTable, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set Table Border width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Border Style settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_TableBorderWidth($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table Border Width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Table's Border Width settings are: " & @CRLF & _
			"The Top Border width is, in Hundredths of a Millimeter (HMM) " & $avSettings[0] & @CRLF & _
			"The Bottom Border width is, in Hundredths of a Millimeter (HMM) " & $avSettings[1] & @CRLF & _
			"The Left Border width is, in Hundredths of a Millimeter (HMM) " & $avSettings[2] & @CRLF & _
			"The Right Border width is, in Hundredths of a Millimeter (HMM) " & $avSettings[3] & @CRLF & _
			"The Vertical Border width is, in Hundredths of a Millimeter (HMM) " & $avSettings[4] & @CRLF & _
			"The Horizontal Border width is, in Hundredths of a Millimeter (HMM) " & $avSettings[5])

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
