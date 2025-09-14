#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oCOM_Error, $oServiceManager
	Local $MyFunc, $ReturnedFunc
	; You don't need to normally set this, as each function already has it set internally. But to speed up the example I'm going to
	; make a short cut to cause a COM error. This will behave the same as any function in this UDF.
	$oCOM_Error = ObjEvent("AutoIt.Error", "__LOImpress_InternalComErrorHandler")
	#forceref $oCOM_Error

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then _ERROR("Error creating Service Manager Object" & " On Line: " & @ScriptLineNumber)

	; Assign my function to a variable to pass to the ComError User Error.
	$MyFunc = _FunctionForErrors

	; Now set the User COM Error function
	; The First Parameter is my User function I want called any time there is a COM Error.
	_LOImpress_ComError_UserFunction($MyFunc)
	If @error Then _ERROR("Error Assigning User COM Error Function.  Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now cause a COM Error, to demonstrate the function.")

	; Create a COM Error by calling a non existent Method.
	$oServiceManager.FakeMethod()

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now I will retrieve the function's name that I set.")

	; Retrieve the currently set User Function.
	$ReturnedFunc = _LOImpress_ComError_UserFunction(Default)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The function's name is: " & FuncName($ReturnedFunc))

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I Will now clear my set function from being called.")

	; Clear any set User Functions.
	_LOImpress_ComError_UserFunction(Null)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will cause another COM Error, to show the function is no longer set.")

	; Create a COM Error by calling a non existent Method.
	$oServiceManager.FakeMethod()
EndFunc

Func _FunctionForErrors($oObjectError)
	MsgBox($MB_OK + $MB_TOPMOST, "COM Error", "A COM Error occurred, here's what we know:" & @CRLF & _
			"Error Number: 0x" & Hex($oObjectError.number, 8) & @CRLF & _
			"Description: " & $oObjectError.windescription & @CRLF & _
			"At line: " & $oObjectError.scriptline & @CRLF & _
			"Source: " & $oObjectError.source & @CRLF & _
			"Description: " & $oObjectError.description & @CRLF & _
			"helpfile: " & $oObjectError.helpfile & @CRLF & _
			"Help content: " & $oObjectError.helpcontent & @CRLF & _
			"LastdllError: " & $oObjectError.lastdllerror & @CRLF & @CRLF & _
			"Some of the above are, as best I know, almost always blank for Libre Office errors.")
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc
