
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $avReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Set the Document's general properties settings to: NewAuthor = "Daniel, Revisions = 8, Editing time = 840 seconds, Apply User Data = True)
	_LOWriter_DocGenProp($oDoc, "Daniel", 8, 840, True)
	If (@error > 0) Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Document's General properties. Return will be an Array in order of function parameters.
	$avReturn = _LOWriter_DocGenProp($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The document's author is: " & $avReturn[0] & @CRLF & _
			"This document has been revised " & $avReturn[1] & " times." & @CRLF & _
			"The total revision time, in seconds is: " & $avReturn[2] & @CRLF & _
			"Are the User-specific settings saved and loaded in this document? True/False: " & $avReturn[3] & @CRLF & @CRLF & _
			"Press Ok and I will clear all the settings.")

	;reset the Document's general properties settings
	_LOWriter_DocGenProp($oDoc, "Someone", Null, Null, Null, True)
	If (@error > 0) Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Document's General properties again.
	$avReturn = _LOWriter_DocGenProp($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The document's new author is: " & $avReturn[0] & @CRLF & _
			"This document has been revised " & $avReturn[1] & " times." & @CRLF & _
			"The total revision time, in seconds is: " & $avReturn[2] & @CRLF & _
			"Are the User-specific settings saved and loaded in this document? True/False: " & $avReturn[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

