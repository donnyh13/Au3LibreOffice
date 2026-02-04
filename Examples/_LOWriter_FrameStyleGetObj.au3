#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame, $oFrameStyle

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert a Frame with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Frame Style to "Labels"
	_LOWriter_FrameStyleCurrent($oDoc, $oFrame, "Labels")
	If @error Then _ERROR($oDoc, "Failed to set the Text Frame style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now retrieve the Labels Frame style object, and modify some of its settings.")

	; Retrieve the "Labels" frame Style object.
	$oFrameStyle = _LOWriter_FrameStyleGetObj($oDoc, "Labels")
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame style object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set "Labels" frame Style background color to $LO_COLOR_RED
	_LOWriter_FrameStyleAreaColor($oFrameStyle, $LO_COLOR_RED)
	If @error Then _ERROR($oDoc, "Failed to set the Frame style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc
