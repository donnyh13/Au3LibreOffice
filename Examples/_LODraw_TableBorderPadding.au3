#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oTable
	Local $avSettings
	Local $iHMM

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Table.
	$oTable = _LODraw_TableInsert($oPage, 5000, 4000, 3, 3)
	If @error Then _ERROR($oDoc, "Failed to insert a Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Border width.
	_LODraw_TableBorderWidth($oTable, $LOD_SHAPE_BORDER_WIDTH_THICK, $LOD_SHAPE_BORDER_WIDTH_THICK, $LOD_SHAPE_BORDER_WIDTH_THICK, $LOD_SHAPE_BORDER_WIDTH_THICK, $LOD_SHAPE_BORDER_WIDTH_THICK, $LOD_SHAPE_BORDER_WIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set Table Border width settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/2 Inch to Hundredths of a Millimeter (HMM).
	$iHMM = _LO_UnitConvert(0.5, $LO_CONVERT_UNIT_INCH_HMM)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Hundredths of a Millimeter (HMM). Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Table padding to 1/2"
	_LODraw_TableBorderPadding($oTable, $iHMM, $iHMM, $iHMM, $iHMM)
	If @error Then _ERROR($oDoc, "Failed to set Table Border Padding settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Table Border Padding settings. Return will be an Array, with values in order of function parameters.
	$avSettings = _LODraw_TableBorderPadding($oTable)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table Border Padding settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Table Border padding values are as follows: " & @CRLF & _
			"The Top border padding is, in in Hundredths of a Millimeter (HMM) " & $avSettings[0] & @CRLF & _
			"The Bottom border padding is, in in Hundredths of a Millimeter (HMM) " & $avSettings[1] & @CRLF & _
			"The Left border padding is, in in Hundredths of a Millimeter (HMM) " & $avSettings[2] & @CRLF & _
			"The Right border padding is, in Hundredths of a Millimeter (HMM) " & $avSettings[3])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
