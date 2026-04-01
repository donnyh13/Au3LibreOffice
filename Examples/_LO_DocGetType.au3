#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"
#include "..\LibreOfficeWriter.au3"
#include "..\LibreOfficeCalc.au3"
#include "..\LibreOfficeBase.au3"

Example()

Func Example()
	Local $oWriterDoc, $oCalcDoc, $oBaseDoc

	; Create a new Calc Document.
	$oCalcDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Writer Document.
	$oWriterDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New, visible, Blank Libre Office Base Document.
	$oBaseDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Document types of the open Documents are as follows (See UDF Constants):" & @CRLF & _
			"$oBaseDoc: " & _LO_DocGetType($oBaseDoc) & @CRLF & _
			"$oCalcDoc: " & _LO_DocGetType($oCalcDoc) & @CRLF & _
			"$oWriterDoc: " & _LO_DocGetType($oWriterDoc))

	; Close the document.
	_LOBase_DocClose($oBaseDoc, False)
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the Calc Document.
	_LOCalc_DocClose($oCalcDoc, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Close the Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the Writer Document.
	_LOWriter_DocClose($oWriterDoc, False)
	If @error Then _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Close the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oBaseDoc, $oCalcDoc, $oWriterDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oBaseDoc) Then _LOBase_DocSubComponentsClose($oBaseDoc)
	If IsObj($oBaseDoc) Then _LOBase_DocClose($oBaseDoc, False)
	If IsObj($oCalcDoc) Then _LOCalc_DocClose($oCalcDoc, False)
	If IsObj($oWriterDoc) Then _LOWriter_DocClose($oWriterDoc, False)
	Exit
EndFunc
