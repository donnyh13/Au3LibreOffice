#include <MsgBoxConstants.au3>
#include <Array.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $asSheets[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Add a new Sheet named "New Sheet" after the first sheet.
	_LOCalc_SheetAdd($oDoc, "New Sheet", 1)
	If @error Then _ERROR("Failed to Create a new Calc Sheet. Error:" & @error & " Extended:" & @extended)

	; Add a new Sheet Autonamed before the first sheet.
	_LOCalc_SheetAdd($oDoc, Null, 0)
	If @error Then _ERROR("Failed to Create a new Calc Sheet. Error:" & @error & " Extended:" & @extended)

	; Retrieve an Array of Sheet names.
	$asSheets = _LOCalc_SheetsGetNames($oDoc)
	If @error Then _ERROR("Failed to Retrieve an array of Sheet names. Error:" & @error & " Extended:" & @extended)

	_ArrayDisplay($asSheets)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR
