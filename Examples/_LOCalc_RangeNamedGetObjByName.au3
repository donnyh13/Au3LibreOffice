#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oNamedRange
	Local $asNamedRanges[0]
	Local $sRanges

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Set the range A1:A5 as a Named Range in the Document (Global) Scope.
	_LOCalc_RangeNamedAdd($oDoc, $oCellRange, "My_Global_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Named Ranges. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell range C3 to E3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "C3", "E3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Set the range C3:E3 as a Named Range for the Sheet (local) scope.
	_LOCalc_RangeNamedAdd($oSheet, $oCellRange, "A_Local_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Named Ranges. Error:" & @error & " Extended:" & @extended)

	; Retrieve a list of Global Named ranges for this document.
	$asNamedRanges = _LOCalc_RangeNamedGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve named Ranges list. Error:" & @error & " Extended:" & @extended)

	; Combine the names into a String.
	For $i = 0 To UBound($asNamedRanges) - 1
		$sRanges &= '"' & $asNamedRanges[$i] & '"' & @CRLF
	Next

	MsgBox($MB_OK, "", "The Global Named Range names currently set for this document are: " & @CRLF & $sRanges & @CRLF & @CRLF & _
			"I will now retrieve the Object for this Named Range and delete it.")

	; Retrieve the Object for the global Named Range "My_Global_Named_Range"
	$oNamedRange = _LOCalc_RangeNamedGetObjByName($oDoc, "My_Global_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to retrieve named Range Object. Error:" & @error & " Extended:" & @extended)

	; Delete the Named Range
	_LOCalc_RangeNamedDelete($oDoc, $oNamedRange)
	If @error Then _ERROR($oDoc, "Failed to delete named Range Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve a list of Global Named ranges for this document.
	$asNamedRanges = _LOCalc_RangeNamedGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve named Ranges list. Error:" & @error & " Extended:" & @extended)

	$sRanges = ""

	; Combine the names into a String.
	For $i = 0 To UBound($asNamedRanges) - 1
		$sRanges &= '"' & $asNamedRanges[$i] & '"' & @CRLF
	Next

	MsgBox($MB_OK, "", "The Global Named Range names currently set for this document are: " & @CRLF & $sRanges & @CRLF & @CRLF & _
			"Now I will demonstrate the same for a local (Sheet) Named Range.")

	; Retrieve a list of Local Named ranges for this Sheet.
	$asNamedRanges = _LOCalc_RangeNamedGetNames($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve named Ranges list. Error:" & @error & " Extended:" & @extended)

	$sRanges = ""

	; Combine the names into a String.
	For $i = 0 To UBound($asNamedRanges) - 1
		$sRanges &= '"' & $asNamedRanges[$i] & '"' & @CRLF
	Next

	MsgBox($MB_OK, "", "The Local Named Range names currently set for this Sheet are: " & @CRLF & $sRanges & @CRLF & @CRLF & _
			"I will now retrieve the Object for this Named Range and delete it.")

	; Retrieve the Object for the local Named Range "A_Local_Named_Range"
	$oNamedRange = _LOCalc_RangeNamedGetObjByName($oSheet, "A_Local_Named_Range")
	If @error Then _ERROR($oDoc, "Failed to retrieve named Range Object. Error:" & @error & " Extended:" & @extended)

	; Delete the Named Range
	_LOCalc_RangeNamedDelete($oSheet, $oNamedRange)
	If @error Then _ERROR($oDoc, "Failed to delete named Range Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve a list of Local Named ranges for this Sheet.
	$asNamedRanges = _LOCalc_RangeNamedGetNames($oSheet)
	If @error Then _ERROR($oDoc, "Failed to retrieve named Ranges list. Error:" & @error & " Extended:" & @extended)

	$sRanges = ""

	; Combine the names into a String.
	For $i = 0 To UBound($asNamedRanges) - 1
		$sRanges &= '"' & $asNamedRanges[$i] & '"' & @CRLF
	Next

	MsgBox($MB_OK, "", "The Local Named Range names currently set for this Sheet are: " & @CRLF & $sRanges)

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc
