#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc
	Local $asLayers
	Local $sLayers = ""

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new layer.
	_LODraw_DocLayerInsert($oDoc, "AutoIt Testing")
	If @error Then _ERROR($oDoc, "Failed to insert a new layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all Layers in the Document.
	$asLayers = _LODraw_DocLayersGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Layer names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Cycle through and list all the Layers.
	For $i = 0 To @extended - 1
		$sLayers &= $asLayers[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "This Document contains the following Layers:" & @CRLF & $sLayers & @CRLF & @CRLF & " Press ok to delete the layer ""AutoIt Testing"".")

	; Delete the layer.
	_LODraw_DocLayerDeleteByName($oDoc, "AutoIt Testing")
	If @error Then _ERROR($oDoc, "Failed to delete a layer. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all Layers in the Document.
	$asLayers = _LODraw_DocLayersGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Layer names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sLayers = ""

	; Cycle through and list all the Layers.
	For $i = 0 To @extended - 1
		$sLayers &= $asLayers[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now the Document contains the following Layers:" & @CRLF & $sLayers)

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
