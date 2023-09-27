
#include "..\LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Success", "A New Writer Document was successfully opened. Press ""OK"" to close it.")

	;Close the document, don't save changes.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then MsgBox($MB_OK, "Error", "Failed to close opened L.O. Document.")

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

