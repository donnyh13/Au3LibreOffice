
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oFrameStyle, $oViewCursor, $oFrame
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new FrameStyle named "Test Style"
	$oFrameStyle = _LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert a Frame with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to insert a Text Frame. Error:" & @error & " Extended:" & @extended)

	;Set the Frame Style to "Test Style"
	_LOWriter_FrameStyleSet($oDoc, $oFrame, "Test Style")
	If (@error > 0) Then _ERROR("Failed to set the Text Frame style. Error:" & @error & " Extended:" & @extended)

	;Set "Test Style" frame Style background color to $LOW_COLOR_RED
	_LOWriter_FrameStyleAreaColor($oFrameStyle, $LOW_COLOR_RED, False)
	If (@error > 0) Then _ERROR("Failed to set the Frame style settings. Error:" & @error & " Extended:" & @extended)

	;See if a Frame Style called "Test Style" exists.
	$bReturn = _LOWriter_FrameStyleExists($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to query for Frame Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a frame style called ""Test Style"" exist for this document? True/False: " & $bReturn)

	;Delete the newly created Frame Style, set Force delete to True, setting the replacement style to "Labels"
	_LOWriter_FrameStyleDelete($oDoc, $oFrameStyle, True, "Labels")
	If (@error > 0) Then _ERROR("Failed to delete a Frame Style. Error:" & @error & " Extended:" & @extended)

	;See if a Frame Style called "Test Style" exists.
	$bReturn = _LOWriter_FrameStyleExists($oDoc, "Test Style")
	If (@error > 0) Then _ERROR("Failed to query for Frame Style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does a frame style called ""Test Style"" exist for this document? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc



