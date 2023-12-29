
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame, $oFrameStyle

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert a Frame with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert a Frame
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If (@error > 0) Then _ERROR("Failed to insert a Text Frame. Error:" & @error & " Extended:" & @extended)

	;Set the Frame Style to "Labels"
	_LOWriter_FrameStyleSet($oDoc, $oFrame, "Labels")
	If (@error > 0) Then _ERROR("Failed to set the Text Frame style. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now retrieve the Labels Frame style object, and modify some of its settings.")

	;Retrieve the "Labels" frame Style object.
	$oFrameStyle = _LOWriter_FrameStyleGetObj($oDoc, "Labels")
	If (@error > 0) Then _ERROR("Failed to retrieve Frame style object. Error:" & @error & " Extended:" & @extended)

	;Set "Labels" frame Style background color to $LOW_COLOR_RED
	_LOWriter_FrameStyleAreaColor($oFrameStyle, $LOW_COLOR_RED, False)
	If (@error > 0) Then _ERROR("Failed to set the Frame style settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

