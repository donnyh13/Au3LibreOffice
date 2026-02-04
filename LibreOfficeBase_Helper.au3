#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Base
#include "LibreOfficeBase_Constants.au3"
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and retrieving data for use in various functions in LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOBase_ComError_UserFunction
; _LOBase_DateStructCreate
; _LOBase_DateStructModify
; _LOBase_FontDescCreate
; _LOBase_FontDescEdit
; _LOBase_FontExists
; _LOBase_FontsGetNames
; _LOBase_FormatKeyCreate
; _LOBase_FormatKeyDelete
; _LOBase_FormatKeyExists
; _LOBase_FormatKeyGetStandard
; _LOBase_FormatKeyGetString
; _LOBase_FormatKeysGetList
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOBase_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default is Default. Accepts a Function, or the Keyword Default and Null. If called with a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $vUserFunction Not a Function, or Default keyword, or Null Keyword.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully set the UserFunction.
;                  @Error 0 @Extended 0 Return 2 = Successfully cleared the set UserFunction.
;                  @Error 0 @Extended 0 Return Function = Returning the set UserFunction.
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
Func _LOBase_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
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
EndFunc   ;==>_LOBase_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DateStructCreate
; Description ...: Create a Date Structure for inserting a Date into certain other functions.
; Syntax ........: _LOBase_DateStructCreate([$iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $iYear               - [optional] an integer value. Default is Null. The Year, as a 4 digit Integer.
;                  $iMonth              - [optional] an integer value (0-12). Default is Null. The Month, as a 2 digit Integer. Call with 0 for Void date.
;                  $iDay                - [optional] an integer value (0-31). Default is Null. The Day, as a 2 digit Integer. Call with 0 for Void date.
;                  $iHours              - [optional] an integer value (0-23). Default is Null. The Hour, as a 2 digit Integer.
;                  $iMinutes            - [optional] an integer value (0-59). Default is Null. Minutes, as a 2 digit Integer.
;                  $iSeconds            - [optional] an integer value (0-59). Default is Null. Seconds, as a 2 digit Integer.
;                  $iNanoSeconds        - [optional] an integer value (0-999,999,999). Default is Null. Nano-Second, as an Integer.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If True: time zone is UTC Else False: unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: Structure.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iYear not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $iYear not 4 digits long.
;                  @Error 1 @Extended 3 Return 0 = $iMonth not an Integer, less than 0 or greater than 12.
;                  @Error 1 @Extended 4 Return 0 = $iDay not an Integer, less than 0 or greater than 31.
;                  @Error 1 @Extended 5 Return 0 = $iHours not an Integer, less than 0 or greater than 23.
;                  @Error 1 @Extended 6 Return 0 = $iMinutes not an Integer, less than 0 or greater than 59.
;                  @Error 1 @Extended 7 Return 0 = $iSeconds not an Integer, less than 0 or greater than 59.
;                  @Error 1 @Extended 8 Return 0 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;                  @Error 1 @Extended 9 Return 0 = $bIsUTC not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.util.DateTime" Object.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Successfully created the Date/Time Structure, Returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling a value with Null keyword will auto fill the value with the current value, such as current hour, etc.
; Related .......: _LOBase_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DateStructCreate($iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tDateStruct

	$tDateStruct = __LO_CreateStruct("com.sun.star.util.DateTime")
	If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$tDateStruct.Year = $iYear

	Else
		$tDateStruct.Year = @YEAR
	EndIf

	If ($iMonth <> Null) Then
		If Not __LO_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tDateStruct.Month = $iMonth

	Else
		$tDateStruct.Month = @MON
	EndIf

	If ($iDay <> Null) Then
		If Not __LO_IntIsBetween($iDay, 0, 31) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tDateStruct.Day = $iDay

	Else
		$tDateStruct.Day = @MDAY
	EndIf

	If ($iHours <> Null) Then
		If Not __LO_IntIsBetween($iHours, 0, 23) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tDateStruct.Hours = $iHours

	Else
		$tDateStruct.Hours = @HOUR
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LO_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tDateStruct.Minutes = $iMinutes

	Else
		$tDateStruct.Minutes = @MIN
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LO_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tDateStruct.Seconds = $iSeconds

	Else
		$tDateStruct.Seconds = @SEC
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LO_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tDateStruct.NanoSeconds = $iNanoSeconds

	Else
		$tDateStruct.NanoSeconds = 0
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LO_VersionCheck(4.1) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$tDateStruct.IsUTC = $bIsUTC

	Else
		If __LO_VersionCheck(4.1) Then $tDateStruct.IsUTC = False
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $tDateStruct)
EndFunc   ;==>_LOBase_DateStructCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_DateStructModify
; Description ...: Set or retrieve Date Structure settings.
; Syntax ........: _LOBase_DateStructModify(ByRef $tDateStruct[, $iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $tDateStruct         - [in/out] a dll struct value. The Date Structure to modify, returned from a _LOBase_DateStructCreate, or setting retrieval function. Structure will be directly modified.
;                  $iYear               - [optional] an integer value. Default is Null. The Year, as a 4 digit Integer.
;                  $iMonth              - [optional] an integer value (0-12). Default is Null. The Month, as a 2 digit Integer. Call with 0 for Void date.
;                  $iDay                - [optional] an integer value (0-31). Default is Null. The Day, as a 2 digit Integer. Call with 0 for Void date.
;                  $iHours              - [optional] an integer value (0-23). Default is Null. The Hour, as a 2 digit Integer.
;                  $iMinutes            - [optional] an integer value (0-59). Default is Null. Minutes, as a 2 digit Integer.
;                  $iSeconds            - [optional] an integer value (0-59). Default is Null. Seconds, as a 2 digit Integer.
;                  $iNanoSeconds        - [optional] an integer value (0-999,999,999). Default is Null. Nano-Second, as an Integer.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If True: time zone is UTC Else False: unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $tDateStruct not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iYear not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iYear not 4 digits long.
;                  @Error 1 @Extended 4 Return 0 = $iMonth not an Integer, less than 0 or greater than 12.
;                  @Error 1 @Extended 5 Return 0 = $iDay not an Integer, less than 0 or greater than 31.
;                  @Error 1 @Extended 6 Return 0 = $iHours not an Integer, less than 0 or greater than 23.
;                  @Error 1 @Extended 7 Return 0 = $iMinutes not an Integer, less than 0 or greater than 59.
;                  @Error 1 @Extended 8 Return 0 = $iSeconds not an Integer, less than 0 or greater than 59.
;                  @Error 1 @Extended 9 Return 0 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;                  @Error 1 @Extended 10 Return 0 = $bIsUTC not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iYear
;                  |                               2 = Error setting $iMonth
;                  |                               4 = Error setting $iDay
;                  |                               8 = Error setting $iHours
;                  |                               16 = Error setting $iMinutes
;                  |                               32 = Error setting $iSeconds
;                  |                               64 = Error setting $iNanoSeconds
;                  |                               128 = Error setting $bIsUTC
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 or 8 Element Array with values in order of function parameters. If current Libre Office version is less than 4.1, the Array will contain 7 elements, as $bIsUTC will be eliminated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOBase_DateStructCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_DateStructModify(ByRef $tDateStruct, $iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avMod[7]

	If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iYear, $iMonth, $iDay, $iHours, $iMinutes, $iSeconds, $iNanoSeconds, $bIsUTC) Then
		If __LO_VersionCheck(4.1) Then
			__LO_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds(), $tDateStruct.IsUTC())

		Else
			__LO_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avMod)
	EndIf

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$tDateStruct.Year = $iYear
		$iError = ($tDateStruct.Year() = $iYear) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iMonth <> Null) Then
		If Not __LO_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tDateStruct.Month = $iMonth
		$iError = ($tDateStruct.Month() = $iMonth) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iDay <> Null) Then
		If Not __LO_IntIsBetween($iDay, 0, 31) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tDateStruct.Day = $iDay
		$iError = ($tDateStruct.Day() = $iDay) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iHours <> Null) Then
		If Not __LO_IntIsBetween($iHours, 0, 23) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tDateStruct.Hours = $iHours
		$iError = ($tDateStruct.Hours() = $iHours) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LO_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tDateStruct.Minutes = $iMinutes
		$iError = ($tDateStruct.Minutes() = $iMinutes) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LO_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tDateStruct.Seconds = $iSeconds
		$iError = ($tDateStruct.Seconds() = $iSeconds) ? ($iError) : (BitOR($iError, 32))
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LO_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tDateStruct.NanoSeconds = $iNanoSeconds
		$iError = ($tDateStruct.NanoSeconds() = $iNanoSeconds) ? ($iError) : (BitOR($iError, 64))
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		If Not __LO_VersionCheck(4.1) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$tDateStruct.IsUTC = $bIsUTC
		$iError = ($tDateStruct.IsUTC() = $bIsUTC) ? ($iError) : (BitOR($iError, 128))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOBase_DateStructModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FontDescCreate
; Description ...: Create a Font Descriptor Map.
; Syntax ........: _LOBase_FontDescCreate([$sFontName = ""[, $iWeight = $LOB_WEIGHT_DONT_KNOW[, $iSlant = $LOB_POSTURE_DONTKNOW[, $nSize = 0[, $iColor = $LO_COLOR_OFF[, $iUnderlineStyle = $LOB_UNDERLINE_DONT_KNOW[, $iUnderlineColor = $LO_COLOR_OFF[, $iStrikelineStyle = $LOB_STRIKEOUT_DONT_KNOW[, $bIndividualWords = False[, $iRelief = $LOB_RELIEF_NONE[, $iCase = $LOB_CASEMAP_NONE[, $bHidden = False[, $bOutline = False[, $bShadow = False]]]]]]]]]]]]]])
; Parameters ....: $sFontName           - [optional] a string value. Default is "". The Font name.
;                  $iWeight             - [optional] an integer value (0-200). Default is $LOB_WEIGHT_DONT_KNOW. The Font weight. See Constants $LOB_WEIGHT_* as defined in LibreOfficeBase_Constants.au3.
;                  $iSlant              - [optional] an integer value (0-5). Default is $LOB_POSTURE_DONTKNOW. The Font italic setting. See Constants $LOB_POSTURE_* as defined in LibreOfficeBase_Constants.au3.
;                  $nSize               - [optional] a general number value. Default is 0. The Font size.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is $LO_COLOR_OFF. The Font Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iUnderlineStyle     - [optional] an integer value (0-18). Default is $LOB_UNDERLINE_DONT_KNOW. The Font underline Style. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeBase_Constants.au3.
;                  $iUnderlineColor     - [optional] an integer value (-1-16777215). Default is $LO_COLOR_OFF. The Font Underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iStrikelineStyle    - [optional] an integer value (0-6). Default is $LOB_STRIKEOUT_DONT_KNOW. The Strikeout line style. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeBase_Constants.au3.
;                  $bIndividualWords    - [optional] a boolean value. Default is False. If True, only individual words are underlined.
;                  $iRelief             - [optional] an integer value (0-2). Default is $LOB_RELIEF_NONE. The Font relief style. See Constants $LOB_RELIEF_* as defined in LibreOfficeBase_Constants.au3.
;                  $iCase               - [optional] an integer value (0-4). Default is $LOB_CASEMAP_NONE. The Character Case Style. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  $bHidden             - [optional] a boolean value. Default is False. If True, the Characters are hidden.
;                  $bOutline            - [optional] a boolean value. Default is False. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is False. If True, the characters have a shadow.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 2 Return 0 = Font called in $sFontName not found.
;                  @Error 1 @Extended 3 Return 0 = $iWeight not an Integer, less than 0 or greater than 200. See Constants $LOB_WEIGHT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iSlant not an Integer, less than 0 or greater than 5. See Constants $LOB_POSTURE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $nSize not a number.
;                  @Error 1 @Extended 6 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 7 Return 0 = $iUnderlineStyle not an Integer, less than 0 or greater than 18. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iUnderlineColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 9 Return 0 = $iStrikelineStyle not an Integer, less than 0 or greater than 6. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 10 Return 0 = $bIndividualWords not a Boolean.
;                  @Error 1 @Extended 11 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants $LOB_RELIEF_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 12 Return 0 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 14 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bShadow not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Returning the created Map Font Descriptor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FontDescCreate($sFontName = "", $iWeight = $LOB_WEIGHT_DONT_KNOW, $iSlant = $LOB_POSTURE_DONTKNOW, $nSize = 0, $iColor = $LO_COLOR_OFF, $iUnderlineStyle = $LOB_UNDERLINE_DONT_KNOW, $iUnderlineColor = $LO_COLOR_OFF, $iStrikelineStyle = $LOB_STRIKEOUT_DONT_KNOW, $bIndividualWords = False, $iRelief = $LOB_RELIEF_NONE, $iCase = $LOB_CASEMAP_NONE, $bHidden = False, $bOutline = False, $bShadow = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $mFontDesc[]

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not _LOBase_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iWeight, $LOB_WEIGHT_DONT_KNOW, $LOB_WEIGHT_BLACK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iSlant, $LOB_POSTURE_NONE, $LOB_POSTURE_REV_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsNumber($nSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not __LO_IntIsBetween($iUnderlineStyle, $LOB_UNDERLINE_NONE, $LOB_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not __LO_IntIsBetween($iUnderlineColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
	If Not __LO_IntIsBetween($iStrikelineStyle, $LOB_STRIKEOUT_NONE, $LOB_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If Not IsBool($bIndividualWords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
	If Not __LO_IntIsBetween($iRelief, $LOB_RELIEF_NONE, $LOB_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
	If Not __LO_IntIsBetween($iCase, $LOB_CASEMAP_NONE, $LOB_CASEMAP_SM_CAPS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)
	If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)
	If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)
	If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

	$mFontDesc.CharFontName = $sFontName
	$mFontDesc.CharWeight = $iWeight
	$mFontDesc.CharPosture = $iSlant
	$mFontDesc.CharHeight = $nSize
	$mFontDesc.CharColor = $iColor
	$mFontDesc.CharUnderline = $iUnderlineStyle
	$mFontDesc.CharUnderlineColor = $iUnderlineColor
	$mFontDesc.CharStrikeout = $iStrikelineStyle
	$mFontDesc.CharWordMode = $bIndividualWords
	$mFontDesc.CharRelief = $iRelief
	$mFontDesc.CharCaseMap = $iCase
	$mFontDesc.CharHidden = $bHidden
	$mFontDesc.CharContoured = $bOutline
	$mFontDesc.CharShadowed = $bShadow

	Return SetError($__LO_STATUS_SUCCESS, 0, $mFontDesc)
EndFunc   ;==>_LOBase_FontDescCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FontDescEdit
; Description ...: Set or Retrieve Font Descriptor settings.
; Syntax ........: _LOBase_FontDescEdit(ByRef $mFontDesc[, $sFontName = Null[, $iWeight = Null[, $iSlant = Null[, $nSize = Null[, $iColor = Null[, $iUnderlineStyle = Null[, $iUnderlineColor = Null[, $iStrikelineStyle = Null[, $bIndividualWords = Null[, $iRelief = Null[, $iCase = Null[, $bHidden = Null[, $bOutline = Null[, $bShadow = Null]]]]]]]]]]]]]])
; Parameters ....: $mFontDesc           - [in/out] a map. A Font descriptor Map as returned from a _LOBase_FontDescCreate, or control property return function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font name.
;                  $iWeight             - [optional] an integer value (0-200). Default is Null. The Font weight. See Constants $LOB_WEIGHT_* as defined in LibreOfficeBase_Constants.au3.
;                  $iSlant              - [optional] an integer value (0-5). Default is Null. The Font italic setting. See Constants $LOB_POSTURE_* as defined in LibreOfficeBase_Constants.au3.
;                  $nSize               - [optional] a general number value. Default is Null. The Font size.
;                  $iColor              - [optional] an integer value (-1-16777215). Default is Null. The Font Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iUnderlineStyle     - [optional] an integer value (0-18). Default is Null. The Font underline Style. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeBase_Constants.au3.
;                  $iUnderlineColor     - [optional] an integer value (-1-16777215). Default is Null.
;                  $iStrikelineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout line style. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeBase_Constants.au3.
;                  $bIndividualWords    - [optional] a boolean value. Default is Null. If True, only individual words are underlined.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Font relief style. See Constants $LOB_RELIEF_* as defined in LibreOfficeBase_Constants.au3.
;                  $iCase               - [optional] an integer value (0-4). Default is Null. The Character Case Style. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, the Characters are hidden.
;                  $bOutline            - [optional] a boolean value. Default is False. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is False. If True, the characters have a shadow.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $mFontDesc not a Map.
;                  @Error 1 @Extended 2 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 3 Return 0 = Font called in $sFontName not found.
;                  @Error 1 @Extended 4 Return 0 = $iWeight not an Integer, less than 0 or greater than 200. See Constants $LOB_WEIGHT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iSlant not an Integer, less than 0 or greater than 5. See Constants $LOB_POSTURE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $nSize not a number.
;                  @Error 1 @Extended 7 Return 0 = $iColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 8 Return 0 = $iUnderlineStyle not an Integer, less than 0 or greater than 18. See Constants $LOB_UNDERLINE_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iUnderlineColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 10 Return 0 = $iStrikelineStyle not an Integer, less than 0 or greater than 6. See Constants $LOB_STRIKEOUT_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 11 Return 0 = $bIndividualWords not a Boolean.
;                  @Error 1 @Extended 12 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants $LOB_RELIEF_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 13 Return 0 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOB_CASEMAP_* as defined in LibreOfficeBase_Constants.au3.
;                  @Error 1 @Extended 14 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 16 Return 0 = $bShadow not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 14 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FontDescEdit(ByRef $mFontDesc, $sFontName = Null, $iWeight = Null, $iSlant = Null, $nSize = Null, $iColor = Null, $iUnderlineStyle = Null, $iUnderlineColor = Null, $iStrikelineStyle = Null, $bIndividualWords = Null, $iRelief = Null, $iCase = Null, $bHidden = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFont[14]

	If Not IsMap($mFontDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sFontName, $iWeight, $iSlant, $nSize, $iColor, $iUnderlineStyle, $iUnderlineColor, $iStrikelineStyle, $bIndividualWords, $iRelief, $iCase, $bHidden, $bOutline, $bShadow) Then
		__LO_ArrayFill($avFont, $mFontDesc.CharFontName, $mFontDesc.CharWeight, $mFontDesc.CharPosture, $mFontDesc.CharHeight, $mFontDesc.CharColor, $mFontDesc.CharUnderline, _
				$mFontDesc.CharUnderlineColor, $mFontDesc.CharStrikeout, $mFontDesc.CharWordMode, $mFontDesc.CharRelief, $mFontDesc.CharCaseMap, $mFontDesc.CharHidden, _
				$mFontDesc.CharContoured, $mFontDesc.CharShadowed)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFont)
	EndIf

	If ($sFontName <> Null) Then
		If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOBase_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$mFontDesc.CharFontName = $sFontName
	EndIf

	If ($iWeight <> Null) Then
		If Not __LO_IntIsBetween($iWeight, $LOB_WEIGHT_DONT_KNOW, $LOB_WEIGHT_BLACK) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$mFontDesc.CharWeight = $iWeight
	EndIf

	If ($iSlant <> Null) Then
		If Not __LO_IntIsBetween($iSlant, $LOB_POSTURE_NONE, $LOB_POSTURE_REV_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$mFontDesc.CharPosture = $iSlant
	EndIf

	If ($nSize <> Null) Then
		If Not IsNumber($nSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$mFontDesc.CharHeight = $nSize
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$mFontDesc.CharColor = $iColor
	EndIf

	If ($iUnderlineStyle <> Null) Then
		If Not __LO_IntIsBetween($iUnderlineStyle, $LOB_UNDERLINE_NONE, $LOB_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$mFontDesc.CharUnderline = $iUnderlineStyle
	EndIf

	If ($iUnderlineColor <> Null) Then
		If Not __LO_IntIsBetween($iUnderlineColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$mFontDesc.CharUnderlineColor = $iUnderlineColor
	EndIf

	If ($iStrikelineStyle <> Null) Then
		If Not __LO_IntIsBetween($iStrikelineStyle, $LOB_STRIKEOUT_NONE, $LOB_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$mFontDesc.CharStrikeout = $iStrikelineStyle
	EndIf

	If ($bIndividualWords <> Null) Then
		If Not IsBool($bIndividualWords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$mFontDesc.CharWordMode = $bIndividualWords
	EndIf

	If ($iRelief <> Null) Then
		If Not __LO_IntIsBetween($iRelief, $LOB_RELIEF_NONE, $LOB_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$mFontDesc.CharRelief = $iRelief
	EndIf

	If ($iCase <> Null) Then
		If Not __LO_IntIsBetween($iCase, $LOB_CASEMAP_NONE, $LOB_CASEMAP_SM_CAPS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$mFontDesc.CharCaseMap = $iCase
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

		$mFontDesc.CharHidden = $bHidden
	EndIf

	If ($bOutline <> Null) Then
		If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)

		$mFontDesc.CharContoured = $bOutline
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)

		$mFontDesc.CharShadowed = $bShadow
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOBase_FontDescEdit

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FontExists
; Description ...: Tests whether a specific font exists by name.
; Syntax ........: _LOBase_FontExists($sFontName[, $oDoc = Null])
; Parameters ....: $sFontName           - a string value. The Font name to search for.
;                  $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen, _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFontName not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a Property Struct.
;                  @Error 2 @Extended 4 Return 0 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Font list.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the Font is available, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function may cause a processor usage spike for a moment or two. If you wish to eliminate this, comment out the current sleep function and place a sleep(10) in its place.
;                  $oDoc is optional, if not called, a Writer Document is created invisibly to perform the check.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FontExists($sFontName, $oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDesktop
	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $bClose = False

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name() = $sFontName Then
			If $bClose Then $oDoc.Close(True)

			Return SetError($__LO_STATUS_SUCCESS, 0, True)
		EndIf

		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOBase_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FontsGetNames
; Description ...: Retrieve an array of currently available fonts.
; Syntax ........: _LOBase_FontsGetNames([$oDoc = Null])
; Parameters ....: $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOBase_ReportConnect, _LOBase_ReportOpen, _LOBase_DocOpen, _LOBase_DocConnect, or _LOBase_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a Property Struct.
;                  @Error 2 @Extended 4 Return 0 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Font list.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a 4 Column Array, @Extended is set to the number of results. See remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, a Writer Document is created invisibly to retrieve the list.
;                  Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold, Italic, etc.
;                  Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for easier processing if required.
;                  From personal tests, Slant only returns 0 or 2. This function may cause a processor usage spike for a moment or two.
;                  The returned array will be as follows:
;                  The first column (Array[1][0]) contains the Font Name.
;                  The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOB_WEIGHT_* as defined in LibreOfficeBase_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOB_POSTURE_* as defined in LibreOfficeBase_Constants.au3.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FontsGetNames($oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts, $atProperties[1]
	Local $asFonts[0][4]
	Local $oServiceManager, $oDesktop
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $bClose = False

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	ReDim $asFonts[UBound($atFonts)][4]

	For $i = 0 To UBound($atFonts) - 1
		$asFonts[$i][0] = $atFonts[$i].Name()
		$asFonts[$i][1] = $atFonts[$i].StyleName()
		$asFonts[$i][2] = $atFonts[$i].Weight
		$asFonts[$i][3] = $atFonts[$i].Slant() ; only 0 or 2?
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOBase_FontsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyCreate
; Description ...: Create a Format Key.
; Syntax ........: _LOBase_FormatKeyCreate(ByRef $oObj, $sFormat)
; Parameters ....: $oObj                - [in/out] an object. A Connection or Document object returned by a previous _LOBase_DatabaseConnectionGet, _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $sFormat             - a string value. The format key String to create.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oObj not a Connection Object and not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $sFormat not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Attempted to Create or Retrieve the Format key, but failed.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Format Key was successfully created, returning Format Key Integer.
;                  @Error 0 @Extended 1 Return Integer = Success. Format Key already existed, returning Format Key Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormatKeyDelete, _LOBase_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyCreate(ByRef $oObj, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.sdbc.Connection") And Not $oObj.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sFormat) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $oObj.supportsService("com.sun.star.sdbc.Connection") Then
		$oFormats = $oObj.Parent.NumberFormatsSupplier.getNumberFormats()

	Else
		$oFormats = $oObj.getNumberFormats()
	EndIf

	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 1, $iFormatKey) ; Format already existed

	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iFormatKey) ; Format created

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Failed to create or retrieve Format
EndFunc   ;==>_LOBase_FormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyDelete
; Description ...: Delete a User-Created Format Key.
; Syntax ........: _LOBase_FormatKeyDelete(ByRef $oObj, $iFormatKey)
; Parameters ....: $oObj                - [in/out] an object. A Connection or Document object returned by a previous _LOBase_DatabaseConnectionGet, _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKey          - an Integer value. The User-Created format Key to delete.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oObj not a Connection Object and not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 4 Return 0 = Format Key called in $iFormatKey not found in Document.
;                  @Error 1 @Extended 5 Return 0 = Format Key called in $iFormatKey not User-Created.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to delete key.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Format Key was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormatKeysGetList, _LOBase_FormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyDelete(ByRef $oObj, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.sdbc.Connection") And Not $oObj.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not _LOBase_FormatKeyExists($oObj, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Key not found.

	If $oObj.supportsService("com.sun.star.sdbc.Connection") Then
		$oFormats = $oObj.Parent.NumberFormatsSupplier.getNumberFormats()

	Else
		$oFormats = $oObj.getNumberFormats()
	EndIf

	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; Key not User Created.

	$oFormats.removeByKey($iFormatKey)

	Return (_LOBase_FormatKeyExists($oObj, $iFormatKey) = False) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))
EndFunc   ;==>_LOBase_FormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyExists
; Description ...: Check if a Document contains a certain Format Key.
; Syntax ........: _LOBase_FormatKeyExists(ByRef $oObj, $iFormatKey[, $iFormatType = $LOB_FORMAT_KEYS_ALL])
; Parameters ....: $oObj                - [in/out] an object. A Connection or Document object returned by a previous _LOBase_DatabaseConnectionGet, _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKey          - an Integer value. The Format Key to look for.
;                  $iFormatType         - [optional] an integer value (0-15881). Default is $LOB_FORMAT_KEYS_ALL. The Format Key type to search in. Values can be BitOr'd together. See Constants, $LOB_FORMAT_KEYS_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oObj not a Connection Object and not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iFormatType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to obtain Array of Date/Time Formats.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Format Key exists in document, True is Returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyExists(ByRef $oObj, $iFormatKey, $iFormatType = $LOB_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys[0]
	Local $tLocale

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.sdbc.Connection") And Not $oObj.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iFormatType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $oObj.supportsService("com.sun.star.sdbc.Connection") Then
		$oFormats = $oObj.Parent.NumberFormatsSupplier.getNumberFormats()

	Else
		$oFormats = $oObj.getNumberFormats()
	EndIf

	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aiFormatKeys = $oFormats.queryKeys($iFormatType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($aiFormatKeys[$i] = $iFormatKey) Then Return SetError($__LO_STATUS_SUCCESS, 0, True) ; Doc does contain format Key
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; Doc does not contain format Key
EndFunc   ;==>_LOBase_FormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyGetStandard
; Description ...: Retrieve the Standard Format for a specific Format Key Type.
; Syntax ........: _LOBase_FormatKeyGetStandard(ByRef $oObj, $iFormatKeyType)
; Parameters ....: $oObj                - [in/out] an object. A Connection or Document object returned by a previous _LOBase_DatabaseConnectionGet, _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKeyType      - an Integer value (1-8196). The Format Key type to retrieve the standard Format for. See Constants $LOB_FORMAT_KEYS_* as defined in LibreOfficeBase_Constants.au3.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oObj not a Connection Object and not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKeyType not an Integer, less than 1 or greater than 8196. See Constants $LOB_FORMAT_KEYS_* as defined in LibreOfficeBase_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.lang.Locale" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Standard Format for the requested Format Key Type.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning the Standard Format for the requested Format Key Type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyGetStandard(ByRef $oObj, $iFormatKeyType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $tLocale
	Local $iStandard

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.sdbc.Connection") And Not $oObj.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iFormatKeyType, $LOB_FORMAT_KEYS_DEFINED, $LOB_FORMAT_KEYS_DURATION) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $oObj.supportsService("com.sun.star.sdbc.Connection") Then
		$oFormats = $oObj.Parent.NumberFormatsSupplier.getNumberFormats()

	Else
		$oFormats = $oObj.getNumberFormats()
	EndIf

	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iStandard = $oFormats.getStandardFormat($iFormatKeyType, $tLocale)
	If Not IsInt($iStandard) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iStandard)
EndFunc   ;==>_LOBase_FormatKeyGetStandard

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeyGetString
; Description ...: Retrieve a Format Key String.
; Syntax ........: _LOBase_FormatKeyGetString(ByRef $oObj, $iFormatKey)
; Parameters ....: $oObj                - [in/out] an object. A Connection or Document object returned by a previous _LOBase_DatabaseConnectionGet, _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $iFormatKey          - an Integer value. The Format Key to retrieve the string for.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oObj not a Connection Object and not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iFormatKey not found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Number Formats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve requested Format Key Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returning Format Key's Format String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOBase_FormatKeysGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeyGetString(ByRef $oObj, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats, $oFormatKey

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.sdbc.Connection") And Not $oObj.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not _LOBase_FormatKeyExists($oObj, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $oObj.supportsService("com.sun.star.sdbc.Connection") Then
		$oFormats = $oObj.Parent.NumberFormatsSupplier.getNumberFormats()

	Else
		$oFormats = $oObj.getNumberFormats()
	EndIf

	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oFormatKey = $oFormats.getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; Key not found.

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFormatKey.FormatString())
EndFunc   ;==>_LOBase_FormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOBase_FormatKeysGetList
; Description ...: Retrieve an Array of Format Keys.
; Syntax ........: _LOBase_FormatKeysGetList(ByRef $oObj[, $bIsUser = False[, $bUserOnly = False[, $iFormatKeyType = $LOB_FORMAT_KEYS_ALL]]])
; Parameters ....: $oObj                - [in/out] an object. A Connection or Document object returned by a previous _LOBase_DatabaseConnectionGet, _LOBase_ReportConnect, or _LOBase_ReportOpen function.
;                  $bIsUser             - [optional] a boolean value. Default is False. If True, Adds a third column to the return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Format Keys are returned.
;                  $iFormatKeyType      - [optional] an integer value (0-15881). Default is $LOB_FORMAT_KEYS_ALL. The Format Key type to retrieve an array of. Values can be BitOr'd together. See Constants, $LOB_FORMAT_KEYS_* as defined in LibreOfficeBase_Constants.au3..
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oObj not a Connection Object and not a Report document opened in Design mode.
;                  @Error 1 @Extended 3 Return 0 = $bIsUser not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iFormatKeyType not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.lang.Locale" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve NumberFormats Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to obtain Array of Format Keys.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a 2 or three column Array, depending on current $bIsUser setting. See remarks. @Extended is set to the number of Keys returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Column One (Array[0][0]) will contain the Format Key Integer,
;                  Column two (Array[0][1]) will contain the Format Key String,
;                  If $bIsUser is called with True, Column Three (Array[0][2]) will contain a Boolean, True if the Format Key is User created, else False.
; Related .......: _LOBase_FormatKeyDelete, _LOBase_FormatKeyGetString, _LOBase_FormatKeyGetStandard
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOBase_FormatKeysGetList(ByRef $oObj, $bIsUser = False, $bUserOnly = False, $iFormatKeyType = $LOB_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOBase_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.sdbc.Connection") And Not $oObj.supportsService("com.sun.star.report.ReportDefinition") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iFormatKeyType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$iColumns = ($bIsUser = True) ? ($iColumns) : (2)

	$tLocale = __LO_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $oObj.supportsService("com.sun.star.sdbc.Connection") Then
		$oFormats = $oObj.Parent.NumberFormatsSupplier.getNumberFormats()

	Else
		$oFormats = $oObj.getNumberFormats()
	EndIf

	If Not IsObj($oFormats) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$aiFormatKeys = $oFormats.queryKeys($iFormatKeyType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ReDim $avFormats[UBound($aiFormatKeys)][$iColumns]

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($bUserOnly = True) Then
			If ($oFormats.getbykey($aiFormatKeys[$i]).UserDefined() = True) Then
				$avFormats[$iCount][0] = $aiFormatKeys[$i]
				$avFormats[$iCount][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
				If ($bIsUser = True) Then $avFormats[$iCount][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
				$iCount += 1
			EndIf

		Else
			$avFormats[$i][0] = $aiFormatKeys[$i]
			$avFormats[$i][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
			If ($bIsUser = True) Then $avFormats[$i][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
		EndIf
		Sleep((IsInt($i / $__LOBCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	If ($bUserOnly = True) Then ReDim $avFormats[$iCount][$iColumns]

	Return SetError($__LO_STATUS_SUCCESS, UBound($avFormats), $avFormats)
EndFunc   ;==>_LOBase_FormatKeysGetList
