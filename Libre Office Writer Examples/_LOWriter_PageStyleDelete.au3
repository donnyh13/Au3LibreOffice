
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $sPageStyleName

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	$sPageStyleName = "NewPageStyle"

	;Create a New Page Style.
	$oPageStyle = _LOWriter_PageStyleCreate($oDoc, $sPageStyleName)
	If (@error > 0) Then _ERROR("Failed to create a new Page Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a Page Style Named """ & $sPageStyleName & """ exist in the document? True/False: " & _
			_LOWriter_PageStyleExists($oDoc, $sPageStyleName))

	;Delete the Page Style
	_LOWriter_PageStyleDelete($oDoc, $oPageStyle)
	If (@error > 0) Then _ERROR("Failed to delete the Page Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now does a Page Style Named """ & $sPageStyleName & """exist in the document? True/False: " & _
			_LOWriter_PageStyleExists($oDoc, $sPageStyleName))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



