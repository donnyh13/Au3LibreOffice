#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oMaster
	Local $bExists

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Add a new Master Page
	$oMaster = _LODraw_PageMasterAdd($oDoc, Null, "Au3 Master")
	If @error Then _ERROR($oDoc, "Failed to Insert a new Master page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; See if the new Master Page exists.
	$bExists = _LODraw_PageMasterExists($oDoc, "Au3 Master")
	If @error Then _ERROR($oDoc, "Failed to query for a Master Page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does the Master Page ""Au3 Master"" exist? True/False: " & $bExists)

	; Delete the Master Page named "Au3 Master"
	_LODraw_PageMasterDeleteByObj($oMaster)
	If @error Then _ERROR($oDoc, "Failed to delete a Master Page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; See if the Master Page exists.
	$bExists = _LODraw_PageMasterExists($oDoc, "Au3 Master")
	If @error Then _ERROR($oDoc, "Failed to query for a Master Page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now does the Master Page ""Au3 Master"" exist? True/False: " & $bExists)

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
