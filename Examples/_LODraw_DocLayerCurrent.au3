#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oLayer, $oCurrentLayer
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current layer Object.
	$oCurrentLayer = _LODraw_DocLayerCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Current Layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current properties of the layer, so I can retrieve the name.
	$avSettings = _LODraw_DocLayerModify($oCurrentLayer)
	If @error Then _ERROR($oDoc, "Failed to retrieve current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current layer is: " & $avSettings[0] & ". " & @CRLF & "Press ok to switch to ""Dimension Lines"".")

	; Retrieve the Dimension lines layer Object.
	$oLayer = _LODraw_DocLayerGetObjByName($oDoc, "measurelines")
	If @error Then _ERROR($oDoc, "Failed to retrieve the layer Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the current layer to Dimension Lines.
	_LODraw_DocLayerCurrent($oDoc, $oLayer)
	If @error Then _ERROR($oDoc, "Failed to set current layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current layer Object.
	$oCurrentLayer = _LODraw_DocLayerCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Current Layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current properties of the layer, so I can retrieve the name.
	$avSettings = _LODraw_DocLayerModify($oCurrentLayer)
	If @error Then _ERROR($oDoc, "Failed to retrieve current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now the current layer is: " & $avSettings[0])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc
