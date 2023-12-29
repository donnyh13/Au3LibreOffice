
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oCOM_Error, $oServiceManager
	Local $aReturn
	Local $MyFunc, $ReturnedFunc
	;You don't need to normally set this, as each function already has it set internally. But to speed up the example I'm going to
	;make a shortcut to cause a COM error. This will behave the same as any function in this UDF.
	$oCOM_Error = ObjEvent("AutoIt.Error", "__LOWriter_InternalComErrorHandler")
	#forceref $oCOM_Error

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then _ERROR("Error creating Service Manager Object")

	;Assign my function to a variable to pass to the ComError User Error.
	$MyFunc = _FunctionForErrors

	;Now set the User COM Error function
	;The First Parameter is my User function I want called any time there is a COM Error.
	;the second function parameter is my first optional Parameter, a String, my second optional Parameter is an integer, my third
	;optional parameter is a boolean, the fourth optional parameter is a String, and the fifth optional parameter  is an integer.
	_LOWriter_ComError_UserFunction($MyFunc, "My First Param", 05, False, "Another String", 100)
	If (@error > 0) Then _ERROR("Error Assigning User COM Error Function.  Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now cause a COM Error, to demonstrate the function.")

	;Create a COM Error by calling a non existent Method.
	$oServiceManager.FakeMethod()

	;I will now set the function again, this time with less Parameters.
	_LOWriter_ComError_UserFunction($MyFunc, "My First Param", 2023, "I have only three Parameters now.")
	If (@error > 0) Then _ERROR("Error Assigning User COM Error Function. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will cause another COM Error, to demonstrate the function.")

	;Create a COM Error by calling a non existent Method.
	$oServiceManager.FakeMethod()

	MsgBox($MB_OK, "", "Now I will retrieve the function's name that I set.")

	;Return the currently set User Function and any Parameters by calling the Default keyword in the first parameter.
	$aReturn = _LOWriter_ComError_UserFunction(Default) ;Since I set parameters, the return will be an Array.

	If Not IsArray($aReturn) Then _ERROR("Error retrieving function Array. Error:" & @error & " Extended:" & @extended)

	;Array will be in order of function parameters. The function will be in the first (zeroth) element.
	$ReturnedFunc = $aReturn[0]

	MsgBox($MB_OK, "", "The function's name is: " & FuncName($ReturnedFunc))

	MsgBox($MB_OK, "", "I Will now clear my set function from being called.")

	;Clear the set User Function be calling it with Null in the first Parameter.
	_LOWriter_ComError_UserFunction(Null)

	MsgBox($MB_OK, "", "I will cause another COM Error, to show the function is no longer set.")

	;Create a COM Error by calling a non existent Method.
	$oServiceManager.FakeMethod()

EndFunc

Func _FunctionForErrors($oObjectError, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)

	MsgBox($MB_OK, "A COM Error occurred, here's what we know:", _
			"Error Number: 0x" & Hex($oObjectError.number, 8) & @CRLF & _
			"Description: " & $oObjectError.windescription & @CRLF & _
			"At line: " & $oObjectError.scriptline & @CRLF & _
			"Source: " & $oObjectError.source & @CRLF & _
			"Description: " & $oObjectError.description & @CRLF & _
			"helpfile: " & $oObjectError.helpfile & @CRLF & _
			"Help content: " & $oObjectError.helpcontent & @CRLF & _
			"LastdllError: " & $oObjectError.lastdllerror & @CRLF & @CRLF & _
			"Some of the above are, as best I know, almost always blank for Libre Office errors." & @CRLF & @CRLF & _
			"The Following User set parameters were also passed: " & @CRLF & _
			"Parameter 1: " & $vParam1 & @CRLF & _
			"Parameter 2: " & $vParam2 & @CRLF & _
			"Parameter 3: " & $vParam3 & @CRLF & _
			"Parameter 4: " & $vParam4 & @CRLF & _
			"Parameter 5: " & $vParam5 & @CRLF & @CRLF & _
			"Your own User function dowsn't need to use any, or all Parameters other than a place for $oObjectError, if you like, " & _
			"its just so the option is there.")

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

