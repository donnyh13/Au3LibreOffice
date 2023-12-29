#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now print the new Writer Document. I suggest turning off your printer so you can cancel " & _
			"the print job without wasting paper.")

	; Print the document, 1 copy, Collate = True, "ALL" Pages, Wait = True, Duplex  = Off
	_LOWriter_DocPrint($oDoc, 1, True, "ALL", True, $LOW_DUPLEX_OFF)
	If @error Then _ERROR("Failed to print the L.O. Document. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have now printed the document and then closed it.")
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc
