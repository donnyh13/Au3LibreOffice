
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $sName

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document's Name.
	$sName = _LOWriter_DocGetName($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Document's name is: " & $sName)

	;Retrieve the full name this time.
	$sName = _LOWriter_DocGetName($oDoc, True)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The document's full name is: " & $sName & @CRLF & @CRLF & _
			"This is the name you would use for Autoit Window functions, such as WinMove etc." & @CRLF & @CRLF & _
			"Does the Document window exist in Autoit's eyes? (0 = False, 1 = True) --  " & WinExists($sName))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

