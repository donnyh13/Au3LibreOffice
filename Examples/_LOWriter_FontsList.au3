
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>
#include <Array.au3>

Example()

Func Example()
	Local $oDoc
	Local $asFonts

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array list of font names
	$asFonts = _LOWriter_FontsList($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve Array of font names. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "There were " & @extended & " fonts found. I will now display the array of the results. The Array will have four " & _
			"columns, " & @CRLF & _
			"-the first column contains the font name, " & @CRLF & _
			"-the second column contains the style name, " & @CRLF & _
			"-the third column contains the Font weight (Bold) value, (see constants)," & @CRLF & _
			"-the fourth column contains the font slant (Italic), (See constants).")

	_ArrayDisplay($asFonts)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



