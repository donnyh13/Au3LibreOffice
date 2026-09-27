#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oSlide, $oTable, $oCell
	Local $avSettings
	Local $iHMM

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

	; Retrieve top left Table Cell Object
	$oCell = _LOImpress_TableCellGetObjByPosition($oTable, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Border width.
	_LOImpress_TableCellBorderWidth($oCell, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK, $LOI_SHAPE_BORDER_WIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set Cell Border width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2 Inch to Hundredths of a Millimeter (HMM).
	$iHMM = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell padding to 1/2"
	_LOImpress_TableCellBorderPadding($oCell, $iHMM, $iHMM, $iHMM, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to set Cell Border Padding settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Border Padding settings. Return will be an Array, with values in order of function parameters.
	$avSettings = _LOImpress_TableCellBorderPadding($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Border Padding settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cell Border padding values are as follows: " & @CRLF & _
			"The Top border padding is, in in Hundredths of a Millimeter (HMM) " & $avSettings[0] & @CRLF & _
			"The Bottom border padding is, in in Hundredths of a Millimeter (HMM) " & $avSettings[1] & @CRLF & _
			"The Left border padding is, in in Hundredths of a Millimeter (HMM) " & $avSettings[2] & @CRLF & _
			"The Right border padding is, in Hundredths of a Millimeter (HMM) " & $avSettings[3])

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
