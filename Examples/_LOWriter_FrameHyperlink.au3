
#include "..\LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $avSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame into the document.
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to create a Frame. Error:" & @error & " Extended:" & @extended)

	;Modify the Frame Hyperlink settings. Set the URL to "https://www.autoitscript.com/site/autoit/", set the name to Autoit©
	_LOWriter_FrameHyperlink($oFrame, "https://www.autoitscript.com/site/autoit/", "Autoit©")
	If (@error > 0) Then _ERROR("Failed to set Frame settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the current Frame Hyperlink settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FrameHyperlink($oFrame)
	If (@error > 0) Then _ERROR("Failed to retrieve Frame settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Frame's Hyperlink settings are as follows: " & @CRLF & _
			"The Hyperlink URL is, (if there is one): " & $avSettings[0] & @CRLF & _
			"The name of the Hyperlink is, (if there is one): " & $avSettings[1] & @CRLF & _
			"The Frame to use when opening the URL is, if this is set, (see UDF constants): " & $avSettings[2] & @CRLF & _
			"Use the server side map? True/False: " & $avSettings[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

