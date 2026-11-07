#include "..\LibreOffice.au3"

If $CmdLine[0] <> 1 Then
	ConsoleWrite("Usage: AutoIt3.exe " & @ScriptName & " <LibreOffice Portable root>" & @CRLF)
	Exit 2
EndIf

Exit Example($CmdLine[1])

Func Example($sPortableRoot)
	_LO_InitializePortable2($sPortableRoot)
	If @error Then Return _ExampleError("_LO_InitializePortable2", @error, @extended)

	Local $oManager = __LO_ServiceManager()
	If Not IsObj($oManager) Then Return _ExampleError("ServiceManager", @error, @extended)
	Local $oPaths = $oManager.createInstance("com.sun.star.util.PathSubstitution")
	If Not IsObj($oPaths) Then Return _ExampleError("PathSubstitution", @error, @extended)
	Local $sProfile = $oPaths.substituteVariables("$(user)", True)
	Local $sExpected = _PathFull($sPortableRoot) & "\Data\settings\user"
	If _LO_PathConvert($sProfile, $LO_PATHCONV_PCPATH_RETURN) <> $sExpected Then
		ConsoleWrite("Unexpected LibreOffice profile: " & $sProfile & @CRLF)
		Return 1
	EndIf
	ConsoleWrite("Portable profile verified: " & $sProfile & @CRLF)

	Local $oCalc = _LOCalc_DocCreate(True, True)
	If @error Then Return _ExampleError("_LOCalc_DocCreate", @error, @extended)
	Local $oCell = $oCalc.getSheets().getByIndex(0).getCellByPosition(0, 0)
	$oCell.setValue(42)
	Local $bCalc = ($oCell.getValue() = 42)
	$oCell = Null
	_LOCalc_DocClose($oCalc, False)
	If @error Then Return _ExampleError("_LOCalc_DocClose", @error, @extended)
	ConsoleWrite("Calc A1 round trip: " & ($bCalc ? "PASS" : "FAIL") & @CRLF)

	Local $oWriter = _LOWriter_DocCreate(True, True)
	If @error Then Return _ExampleError("_LOWriter_DocCreate", @error, @extended)
	Local $oText = $oWriter.getText()
	$oText.setString("Portable Writer smoke test")
	Local $bWriter = ($oText.getString() == "Portable Writer smoke test")
	$oText = Null
	_LOWriter_DocClose($oWriter, False)
	If @error Then Return _ExampleError("_LOWriter_DocClose", @error, @extended)
	ConsoleWrite("Writer text round trip: " & ($bWriter ? "PASS" : "FAIL") & @CRLF)

	Local $bTerminated = _LO_Terminate()
	If @error Or Not $bTerminated Then Return _ExampleError("_LO_Terminate", @error, @extended)
	Return ($bCalc And $bWriter) ? 0 : 1
EndFunc   ;==>Example

Func _ExampleError($sStep, $iError, $iExtended)
	ConsoleWrite($sStep & " failed. @error=" & $iError & " @extended=" & $iExtended & @CRLF)
	Return 1
EndFunc   ;==>_ExampleError
