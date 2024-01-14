#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $sPageStyleName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	$sPageStyleName = "NewPageStyle"

	; Create a New Page Style.
	$oPageStyle = _LOWriter_PageStyleCreate($oDoc, $sPageStyleName)
	If @error Then _ERROR($oDoc, "Failed to create a new Page Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Page Style Named """ & $sPageStyleName & """ exist in the document? True/False: " & _
			_LOWriter_PageStyleExists($oDoc, $sPageStyleName))

	; Delete the Page Style
	_LOWriter_PageStyleDelete($oDoc, $oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to delete the Page Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now does a Page Style Named """ & $sPageStyleName & """exist in the document? True/False: " & _
			_LOWriter_PageStyleExists($oDoc, $sPageStyleName))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR
