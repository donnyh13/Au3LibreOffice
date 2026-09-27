#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Impress
#include "LibreOfficeImpress_Constants.au3"
#include "LibreOfficeImpress_Internal.au3"

; Other includes for Impress

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and retrieving data for use in various functions in LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_ComError_UserFunction
; _LOImpress_DateStructCreate
; _LOImpress_DateStructModify
; _LOImpress_FontExists
; _LOImpress_FontsGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOImpress_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
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
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in Autoit HelpFile.) Using the first parameter in the UserFunction.
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
Func _LOImpress_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
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
EndFunc   ;==>_LOImpress_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DateStructCreate
; Description ...: Create a Date Structure for inserting a Date into certain other functions.
; Syntax ........: _LOImpress_DateStructCreate([$iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $iYear               - [optional] Default is Null. The Year, as a 4 digit Integer.
;                  $iMonth              - [optional] (0-12) Default is Null. The Month, as a 2 digit Integer. Call with 0 for Void date.
;                  $iDay                - [optional] (0-31) Default is Null. The Day, as a 2 digit Integer. Call with 0 for Void date.
;                  $iHours              - [optional] (0-23) Default is Null. The Hour, as a 2 digit Integer.
;                  $iMinutes            - [optional] (0-59) Default is Null. Minutes, as a 2 digit Integer.
;                  $iSeconds            - [optional] (0-59) Default is Null. Seconds, as a 2 digit Integer.
;                  $iNanoSeconds        - [optional] (0-999999999) Default is Null. Nano-Second, as an Integer.
;                  $bIsUTC              - [optional] Default is Null. If True: time zone is UTC Else False: unknown time zone. LibreOffice version 4.1 and up.
; Return values .: Success: Structure.
;                  @Error: 0, @Extended: 0, Return: Structure = Success. Successfully created the Date/Time Structure, Returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $iYear not an Integer.
;                  @Error: 1, @Extended: 2 = $iYear not 4 digits long.
;                  @Error: 1, @Extended: 3 = $iMonth not an Integer, less than 0 or greater than 12.
;                  @Error: 1, @Extended: 4 = $iDay not an Integer, less than 0 or greater than 31.
;                  @Error: 1, @Extended: 5 = $iHours not an Integer, less than 0 or greater than 23.
;                  @Error: 1, @Extended: 6 = $iMinutes not an Integer, less than 0 or greater than 59.
;                  @Error: 1, @Extended: 7 = $iSeconds not an Integer, less than 0 or greater than 59.
;                  @Error: 1, @Extended: 8 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;                  @Error: 1, @Extended: 9 = $bIsUTC not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create "com.sun.star.util.DateTime" Object.
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.1.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Calling a value with Null keyword will auto fill the value with the current value, such as current hour, etc.
; Related .......: _LOImpress_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DateStructCreate($iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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
EndFunc   ;==>_LOImpress_DateStructCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DateStructModify
; Description ...: Set or retrieve Date Structure settings.
; Syntax ........: _LOImpress_DateStructModify(ByRef $tDateStruct[, $iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $tDateStruct         - The Date Structure to modify, returned from a _LOImpress_DateStructCreate, or setting retrieval function. Structure will be directly modified.
;                  $iYear               - [optional] Default is Null. The Year, as a 4 digit Integer.
;                  $iMonth              - [optional] (0-12) Default is Null. The Month, as a 2 digit Integer. Call with 0 for Void date.
;                  $iDay                - [optional] (0-31) Default is Null. The Day, as a 2 digit Integer. Call with 0 for Void date.
;                  $iHours              - [optional] (0-23) Default is Null. The Hour, as a 2 digit Integer.
;                  $iMinutes            - [optional] (0-59) Default is Null. Minutes, as a 2 digit Integer.
;                  $iSeconds            - [optional] (0-59) Default is Null. Seconds, as a 2 digit Integer.
;                  $iNanoSeconds        - [optional] (0-999999999) Default is Null. Nano-Second, as an Integer.
;                  $bIsUTC              - [optional] Default is Null. If True: time zone is UTC Else False: unknown time zone. LibreOffice version 4.1 and up.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in an 8 Element Array with values in order of function parameters. If current LibreOffice version is less than 4.1, the $bIsUTC parameter will return a Null value.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $tDateStruct not an Object.
;                  @Error: 1, @Extended: 2 = $iYear not an Integer.
;                  @Error: 1, @Extended: 3 = $iYear not 4 digits long.
;                  @Error: 1, @Extended: 4 = $iMonth not an Integer, less than 0 or greater than 12.
;                  @Error: 1, @Extended: 5 = $iDay not an Integer, less than 0 or greater than 31.
;                  @Error: 1, @Extended: 6 = $iHours not an Integer, less than 0 or greater than 23.
;                  @Error: 1, @Extended: 7 = $iMinutes not an Integer, less than 0 or greater than 59.
;                  @Error: 1, @Extended: 8 = $iSeconds not an Integer, less than 0 or greater than 59.
;                  @Error: 1, @Extended: 9 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;                  @Error: 1, @Extended: 10 = $bIsUTC not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iYear
;                  |                               2 = Error setting $iMonth
;                  |                               4 = Error setting $iDay
;                  |                               8 = Error setting $iHours
;                  |                               16 = Error setting $iMinutes
;                  |                               32 = Error setting $iSeconds
;                  |                               64 = Error setting $iNanoSeconds
;                  |                               128 = Error setting $bIsUTC
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.1.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_DateStructCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DateStructModify(ByRef $tDateStruct, $iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avMod[8]

	If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iYear, $iMonth, $iDay, $iHours, $iMinutes, $iSeconds, $iNanoSeconds, $bIsUTC) Then
		If __LO_VersionCheck(4.1) Then
			__LO_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds(), $tDateStruct.IsUTC())

		Else
			__LO_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds(), Null)
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
EndFunc   ;==>_LOImpress_DateStructModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FontExists
; Description ...: Tests whether a specific font exists by name.
; Syntax ........: _LOImpress_FontExists($sFontName[, $oDoc = Null])
; Parameters ....: $sFontName           - The Font name to search for.
;                  $oDoc                - [optional] Default is Null. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Returning True if the Font is available, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $sFontName not a String.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error: 2, @Extended: 3 = Failed to create a Property Struct.
;                  @Error: 2, @Extended: 4 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Font list.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, an Impress Document is created invisibly to perform the check.
; Related .......: _LOImpress_FontsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FontExists($sFontName, $oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/simpress", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name = $sFontName Then
			If $bClose Then $oDoc.Close(True)

			Return SetError($__LO_STATUS_SUCCESS, 0, True)
		EndIf
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOImpress_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FontsGetNames
; Description ...: Retrieve an array of currently available font names.
; Syntax ........: _LOImpress_FontsGetNames([$oDoc = Null])
; Parameters ....: $oDoc                - [optional] Default is Null. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning a 4 Column Array, @Extended is set to the number of results. See remarks
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error: 2, @Extended: 3 = Failed to create a Property Struct.
;                  @Error: 2, @Extended: 4 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Font list.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, an Impress Document is created invisibly to perform the check.
;                  Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold, Italic, etc. Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for easier processing if required.
;                  From personal tests, Slant only returns 0 or 2.
;                  The returned array will be as follows:
;                  The first column (Array[1][0]) contains the Font Name.
;                  The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3.
; Related .......: _LOImpress_FontExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FontsGetNames($oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asFonts[0][4]
	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/simpress", "_blank", $iURLFrameCreate, $atProperties)
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
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOImpress_FontsGetNames
