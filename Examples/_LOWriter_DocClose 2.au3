
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $sSaveName, $sSavepath
	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Success", "A New Writer Document was successfully opened. Press OK to close and save it.")

	;Create a Temporary Unique File name.
	$sSaveName = "TestCloseDocument_" & @YEAR & "_" & @MON & "_" & @YDAY & "_" & @HOUR & "_" & @MIN & "_" & @SEC

	;Close the document, save changes.
	$sSavepath = _LOWriter_DocClose($oDoc, True, $sSaveName)
	If (@error > 0) Then _ERROR("Failed to close and save opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "Success", "The Writer Document was successfully saved to the following path: " & $sSavepath & @CRLF & _
			"Press OK to Delete it.")

	FileDelete($sSavepath)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



