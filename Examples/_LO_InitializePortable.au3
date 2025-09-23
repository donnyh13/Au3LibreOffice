#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"
#include "..\LibreOfficeWriter.au3"
#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $sVersionAndName
	Local $oWriterDoc, $oCalcDoc

	; Update this path to match the path to your Portable LibreOffice/OpenOffice folder.
	Local $sPathToPortable = "C:\Portable Apps\LibreOfficePortablePrevious"

	; Initialize Portable LibreOffice
	_LO_InitializePortable($sPathToPortable)
	If @error Then _ERROR($oCalcDoc, $oWriterDoc, "Failed to Initialize L.O. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current full Office version number and name.
	$sVersionAndName = _LO_VersionGet(False, True)
	If @error Then _ERROR($oCalcDoc, $oWriterDoc, "Failed to retrieve L.O. version information. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Your current full LibreOffice version, including the name is: " & $sVersionAndName & @CRLF & @CRLF & _
			"Press ok to open a couple new Documents.")

	; Create a new Calc Document.
	$oCalcDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oCalcDoc, $oWriterDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Writer Document.
	$oWriterDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oCalcDoc, $oWriterDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "We have now created several Documents using Portable LibreOffice. Press ok to close them.")

	; Close the Calc Document.
	_LOCalc_DocClose($oCalcDoc, False)
	If @error Then _ERROR($oCalcDoc, $oWriterDoc, "Failed to Close the Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the Writer Document.
	_LOWriter_DocClose($oWriterDoc, False)
	If @error Then _ERROR($oCalcDoc, $oWriterDoc, "Failed to Close the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDocCalc, $oDocWriter, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDocCalc) Then _LOCalc_DocClose($oDocCalc, False)
	If IsObj($oDocWriter) Then _LOWriter_DocClose($oDocWriter, False)
	Exit
EndFunc
