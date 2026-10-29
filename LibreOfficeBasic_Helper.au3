#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Basic
#include "LibreOfficeBasic_Internal.au3"

; Other includes for Basic

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and retrieving data for use in various functions in LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBasic_ComError_UserFunction
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBasic_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOBasic_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] Default is Default. Accepts a Function, or the Keyword Default and Null. If called with a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;                  @Error: 0, @Extended: 0, Return: 1 = Successfully set the UserFunction.
;                  @Error: 0, @Extended: 0, Return: 2 = Successfully cleared the set UserFunction.
;                  @Error: 0, @Extended: 0, Return: Function = Returning the set UserFunction.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $vUserFunction Not a Function, or Default keyword, or Null Keyword.
; Author ........: mLipok
; Modified ......: donnyh13 - Added a clear UserFunction without error option. Also added parameters option.
; Remarks .......: The first parameter passed to the User function will always be the COM Error object. See below.
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in AutoIt Help File.) Using the first parameter in the UserFunction.
;                  For Example MyFunc($oMyError)
;                  - $oMyError.number The Windows HRESULT value from a COM call
;                  - $oMyError.windescription The FormatWinError() text derived from .number
;                  - $oMyError.source Name of the Object generating the error (contents from ExcepInfo.source)
;                  - $oMyError.description Source Object's description of the error (contents from ExcepInfo.description)
;                  - $oMyError.helpfile Source Object's help file for the error (contents from ExcepInfo.helpfile)
;                  - $oMyError.helpcontext Source Object's help file context id number (contents from ExcepInfo.helpcontext)
;                  - $oMyError.lastdllerror The number returned from GetLastError()
;                  - $oMyError.scriptline The script line on which the error was generated
;                  - NOTE: Not all properties will necessarily contain data, some will be blank.
;                  If MsgBox or ConsoleWrite functions are passed to this function, the error details will be displayed using that function automatically.
;                  If called with Default keyword, the current UserFunction, if set, will be returned.
;                  If called with Null keyword, the currently set UserFunction is cleared and only the internal ComErrorHandler will be called for COM Errors.
;                  The stored UserFunction (besides MsgBox and ConsoleWrite) will be called as follows: UserFunc($oComError,$vParam1,$vParam2,$vParam3,$vParam4,$vParam5)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBasic_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
	#forceref $vParam1, $vParam2, $vParam3, $vParam4, $vParam5

	; If user does not set a function, UDF must use internal function to avoid AutoItError.
	Local Static $vUserFunction_Static = Default
	Local $avUserFuncWParams[@NumParams]

	If $vUserFunction = Default Then
		; just return stored static User Function variable

		Return SetError($__LO_STATUS_SUCCESS, 0, $vUserFunction_Static)

	ElseIf IsFunc($vUserFunction) Then
		; If User called Parameters, then add to array.
		If @NumParams > 1 Then
			$avUserFuncWParams[0] = $vUserFunction
			For $i = 1 To @NumParams - 1
				$avUserFuncWParams[$i] = Eval("vParam" & $i)
				; set static variable
			Next
			$vUserFunction_Static = $avUserFuncWParams

		Else
			$vUserFunction_Static = $vUserFunction
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	ElseIf $vUserFunction = Null Then
		; Clear User Function.
		$vUserFunction_Static = Default

		Return SetError($__LO_STATUS_SUCCESS, 0, 2)

	Else
		; return error as an incorrect parameter was passed to this function

		Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOBasic_ComError_UserFunction
