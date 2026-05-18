#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1

#include-once

#include "LibreOffice_Constants.au3"
#include "LibreOffice_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Helper functions for using this UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LO_ComError_UserFunction
; _LO_ConvertColorFromLong
; _LO_ConvertColorToLong
; _LO_DocConnect
; _LO_DocGetType
; _LO_GradientMulticolorAdd
; _LO_GradientMulticolorDelete
; _LO_GradientMulticolorModify
; _LO_InitializePortable
; _LO_PathConvert
; _LO_PrintersGetNames
; _LO_PrintersGetNamesAlt
; _LO_Terminate
; _LO_TransparencyGradientMultiAdd
; _LO_TransparencyGradientMultiDelete
; _LO_TransparencyGradientMultiModify
; _LO_UnitConvert
; _LO_VersionGet
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LO_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default is Default. Accepts a Function, or the Keyword Default and Null. If called with a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;                  @Error: 0, @Extended: 0, Return: 1 = Successfully set the UserFunction.
;                  @Error: 0, @Extended: 0, Return: 2 = Successfully cleared the set UserFunction.
;                  @Error: 0, @Extended: 0, Return: Function = Returning the set UserFunction.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
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
Func _LO_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
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
EndFunc   ;==>_LO_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_ConvertColorFromLong
; Description ...: Convert a RGB Color Integer to Hex, RGB, HSB or CMYK.
; Syntax ........: _LO_ConvertColorFromLong([$iHex = Null[, $iRGB = Null[, $iHSB = Null[, $iCMYK = Null]]]])
; Parameters ....: $iHex                - [optional] Default is Null. Convert a RGB Color Integer to Hexadecimal.
;                  $iRGB                - [optional] Default is Null. Convert a RGB Color Integer to R.G.B.
;                  $iHSB                - [optional] Default is Null. Convert a RGB Color Integer to H.S.B.
;                  $iCMYK               - [optional] Default is Null. Convert a RGB Color Integer to C.M.Y.K.
; Return values .: Success: String or Array.
;                  @Error: 0, @Extended: 1, Return: String = RGB Integer converted To Hexadecimal (as a String). (Without the "0x" prefix)
;                  @Error: 0, @Extended: 2, Return: Array = Array containing RGB Integer converted To Red, Green, Blue,(RGB). $Array[0] = R, $Array[1] = G, etc.
;                  @Error: 0, @Extended: 3, Return: Array = Array containing RGB Integer converted To Hue, Saturation, Brightness, (HSB). $Array[0] = H, $Array[1] = S, etc.
;                  @Error: 0, @Extended: 4, Return: Array = Array containing RGB Integer converted To Cyan, Magenta, Yellow, Black, (CMYK). $Array[0] = C, $Array[1] = M, etc.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = No parameters set.
;                  @Error: 1, @Extended: 2 = No parameters called with an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve a Hexadecimal color value, call the RGB Color Integer in $iHex, To retrieve a R(ed)G(reen)B(lue) color value, call Null in $iHex, and call the RGB Color Integer into $iRGB, etc. for the other color types.
;                  Hex returns as a string variable, all others (RGB, HSB, CMYK) return an array.
;                  The Hexadecimal figure returned doesn't contain the usual "0x", as LibreOffice does not implement it in its numbering system.
; Related .......: _LO_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_ConvertColorFromLong($iHex = Null, $iRGB = Null, $iHSB = Null, $iCMYK = Null)
	Local $nRed, $nGreen, $nBlue, $nResult, $nMaxRGB, $nMinRGB, $nHue, $nSaturation, $nBrightness, $nCyan, $nMagenta, $nYellow, $nBlack
	Local $dHex
	Local $aiReturn[0]

	If (@NumParams = 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Select
		Case IsInt($iHex) ; Long TO Hex
			$nRed = BitAND(BitShift($iHex, 16), 0xff)
			$nGreen = BitAND(BitShift($iHex, 8), 0xff)
			$nBlue = BitAND($iHex, 0xff)

			$dHex = Hex($nRed, 2) & Hex($nGreen, 2) & Hex($nBlue, 2)

			Return SetError($__LO_STATUS_SUCCESS, 1, $dHex)

		Case IsInt($iRGB) ; Long to RGB
			$nRed = BitAND(BitShift($iRGB, 16), 0xff)
			$nGreen = BitAND(BitShift($iRGB, 8), 0xff)
			$nBlue = BitAND($iRGB, 0xff)

			ReDim $aiReturn[3]
			$aiReturn[0] = $nRed
			$aiReturn[1] = $nGreen
			$aiReturn[2] = $nBlue

			Return SetError($__LO_STATUS_SUCCESS, 2, $aiReturn)

		Case IsInt($iHSB) ; Long to HSB
			$nRed = (Mod(($iHSB / 65536), 256) / 255)
			$nGreen = (Mod(($iHSB / 256), 256) / 255)
			$nBlue = (Mod($iHSB, 256) / 255)

			; get Max RGB Value
			$nResult = ($nRed > $nGreen) ? ($nRed) : ($nGreen)
			$nMaxRGB = ($nResult > $nBlue) ? ($nResult) : ($nBlue)
			; get Min RGB Value
			$nResult = ($nRed < $nGreen) ? ($nRed) : ($nGreen)
			$nMinRGB = ($nResult < $nBlue) ? ($nResult) : ($nBlue)

			; Determine Brightness
			$nBrightness = $nMaxRGB
			; Determine Hue
			$nHue = 0
			Select
				Case $nRed = $nGreen = $nBlue ; Red, Green, and Blue are equal.
					$nHue = 0

				Case ($nRed >= $nGreen) And ($nGreen >= $nBlue) ; Red Highest, Blue Lowest
					$nHue = (60 * (($nGreen - $nBlue) / ($nRed - $nBlue)))

				Case ($nRed >= $nBlue) And ($nBlue >= $nGreen) ; Red Highest, Green Lowest
					$nHue = (60 * (6 - (($nBlue - $nGreen) / ($nRed - $nGreen))))

				Case ($nGreen >= $nRed) And ($nRed >= $nBlue) ; Green Highest, Blue Lowest
					$nHue = (60 * (2 - (($nRed - $nBlue) / ($nGreen - $nBlue))))

				Case ($nGreen >= $nBlue) And ($nBlue >= $nRed) ; Green Highest, Red Lowest
					$nHue = (60 * (2 + (($nBlue - $nRed) / ($nGreen - $nRed))))

				Case ($nBlue >= $nGreen) And ($nGreen >= $nRed) ; Blue Highest, Red Lowest
					$nHue = (60 * (4 - (($nGreen - $nRed) / ($nBlue - $nRed))))

				Case ($nBlue >= $nRed) And ($nRed >= $nGreen) ; Blue Highest, Green Lowest
					$nHue = (60 * (4 + (($nRed - $nGreen) / ($nBlue - $nGreen))))
			EndSelect

			; Determine Saturation
			$nSaturation = ($nMaxRGB = 0) ? (0) : (($nMaxRGB - $nMinRGB) / $nMaxRGB)

			$nHue = ($nHue > 0) ? (Round($nHue)) : (0)
			$nSaturation = Round(($nSaturation * 100))
			$nBrightness = Round(($nBrightness * 100))

			ReDim $aiReturn[3]
			$aiReturn[0] = $nHue
			$aiReturn[1] = $nSaturation
			$aiReturn[2] = $nBrightness

			Return SetError($__LO_STATUS_SUCCESS, 3, $aiReturn)

		Case IsInt($iCMYK) ; Long to CMYK
			$nRed = (Mod(($iCMYK / 65536), 256))
			$nGreen = (Mod(($iCMYK / 256), 256))
			$nBlue = (Mod($iCMYK, 256))

			$nRed = Round(($nRed / 255), 3)
			$nGreen = Round(($nGreen / 255), 3)
			$nBlue = Round(($nBlue / 255), 3)

			; get Max RGB Value
			$nResult = ($nRed > $nGreen) ? ($nRed) : ($nGreen)
			$nMaxRGB = ($nResult > $nBlue) ? ($nResult) : ($nBlue)

			$nBlack = (1 - $nMaxRGB)
			$nCyan = ((1 - $nRed - $nBlack) / (1 - $nBlack))
			$nMagenta = ((1 - $nGreen - $nBlack) / (1 - $nBlack))
			$nYellow = ((1 - $nBlue - $nBlack) / (1 - $nBlack))

			$nCyan = Round(($nCyan * 100))
			$nMagenta = Round(($nMagenta * 100))
			$nYellow = Round(($nYellow * 100))
			$nBlack = Round(($nBlack * 100))

			ReDim $aiReturn[4]
			$aiReturn[0] = $nCyan
			$aiReturn[1] = $nMagenta
			$aiReturn[2] = $nYellow
			$aiReturn[3] = $nBlack

			Return SetError($__LO_STATUS_SUCCESS, 4, $aiReturn)

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; no parameters set to an integer
	EndSelect
EndFunc   ;==>_LO_ConvertColorFromLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_ConvertColorToLong
; Description ...: Convert Hex, RGB, HSB or CMYK to a RGB Color Integer.
; Syntax ........: _LO_ConvertColorToLong([$vVal1 = Null[, $vVal2 = Null[, $vVal3 = Null[, $vVal4 = Null]]]])
; Parameters ....: $vVal1               - [optional] Default is Null. See remarks.
;                  $vVal2               - [optional] Default is Null. See remarks.
;                  $vVal3               - [optional] Default is Null. See remarks.
;                  $vVal4               - [optional] Default is Null. See remarks.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 1, Return: Integer = RGB Color Integer converted from Hexadecimal.
;                  @Error: 0, @Extended: 2, Return: Integer = RGB Color Integer converted from Red, Green, Blue, (RGB).
;                  @Error: 0, @Extended: 3, Return: Integer = RGB Color Integer converted from (H)ue, (S)aturation, (B)rightness,
;                  @Error: 0, @Extended: 4, Return: Integer = RGB Color Integer converted from (C)yan, (M)agenta, (Y)ellow, Blac(k)
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = No parameters set.
;                  @Error: 1, @Extended: 2 = One parameter called, but not in String format(Hex).
;                  @Error: 1, @Extended: 3 = Hex parameter contains non Hex characters.
;                  @Error: 1, @Extended: 4 = Hex parameter not 6 characters long.
;                  @Error: 1, @Extended: 5 = Hue parameter contains more than just digits.
;                  @Error: 1, @Extended: 6 = Saturation parameter contains more than just digits.
;                  @Error: 1, @Extended: 7 = Brightness parameter contains more than just digits.
;                  @Error: 1, @Extended: 8 = Three parameters called but not all Integers (RGB) and not all Strings (HSB).
;                  @Error: 1, @Extended: 9 = Four parameters called but not all Integers(CMYK).
;                  @Error: 1, @Extended: 10 = Too many or too few parameters called.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To Convert a Hex(adecimal) color code, call the Hex code in $vVal1 in String Format.
;                  To convert a R(ed) G(reen) B(lue color, call R value in $vVal1 as an Integer, G in $vVal2 as an Integer, and B in $vVal3 as an Integer.
;                  To convert a H(ue) S(aturation) B(rightness) color, call H in $vVal1 as a String, S in $vVal2 as a String, and B in $vVal3 as a string.
;                  To convert C(yan) M(agenta) Y(ellow) Blac(k) call C in $vVal1 as an Integer, M in $vVal2 as an Integer, Y in $vVal3 as an Integer, and K in $vVal4 as an Integer.
;                  The Hexadecimal figure entered cannot contain the usual "0x", as LibreOffice does not implement it in its numbering system.
; Related .......: _LO_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_ConvertColorToLong($vVal1 = Null, $vVal2 = Null, $vVal3 = Null, $vVal4 = Null) ; RGB = Int, CMYK = Int, HSB = String, Hex = String.
	Local Const $__STR_STRIPALL = 8
	Local $iRed, $iGreen, $iBlue, $iLong, $iHue, $iSaturation, $iBrightness
	Local $dHex
	Local $nMaxRGB, $nMinRGB, $nChroma, $nHuePre, $nCyan, $nMagenta, $nYellow, $nBlack

	If (@NumParams = 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch @NumParams
		Case 1 ; Hex
			If Not IsString($vVal1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; not a string

			$vVal1 = StringStripWS($vVal1, $__STR_STRIPALL)
			$dHex = $vVal1

			; From Hex to RGB
			If (StringLen($dHex) = 6) Then
				If StringRegExp($dHex, "[^0-9a-fA-F]") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; $dHex contains non Hex characters.

				$iRed = BitAND(BitShift("0x" & $dHex, 16), 0xFF)
				$iGreen = BitAND(BitShift("0x" & $dHex, 8), 0xFF)
				$iBlue = BitAND("0x" & $dHex, 0xFF)

				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

				Return SetError($__LO_STATUS_SUCCESS, 1, $iLong) ; Long from Hex

			Else

				Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Wrong length of string.
			EndIf

		Case 3 ; RGB and HSB; HSB is all strings, RGB all Integers.
			If (IsInt($vVal1) And IsInt($vVal2) And IsInt($vVal3)) Then ; RGB
				$iRed = $vVal1
				$iGreen = $vVal2
				$iBlue = $vVal3

				; RGB to Long
				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

				Return SetError($__LO_STATUS_SUCCESS, 2, $iLong) ; Long from RGB

			ElseIf IsString($vVal1) And IsString($vVal2) And IsString($vVal3) Then ; Hue Saturation and Brightness (HSB)
				; HSB to RGB
				$vVal1 = StringStripWS($vVal1, $__STR_STRIPALL)
				$vVal2 = StringStripWS($vVal2, $__STR_STRIPALL)
				$vVal3 = StringStripWS($vVal3, $__STR_STRIPALL) ; Strip WS so I can check string length in HSB conversion.

				$iHue = Number($vVal1)
				If (StringLen($vVal1)) <> (StringLen($iHue)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; String contained more than just digits

				$iSaturation = Number($vVal2)
				If (StringLen($vVal2)) <> (StringLen($iSaturation)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; String contained more than just digits

				$iBrightness = Number($vVal3)
				If (StringLen($vVal3)) <> (StringLen($iBrightness)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0) ; String contained more than just digits

				$nMaxRGB = ($iBrightness / 100)
				$nChroma = (($iSaturation / 100) * ($iBrightness / 100))
				$nMinRGB = ($nMaxRGB - $nChroma)
				$nHuePre = ($iHue >= 300) ? (($iHue - 360) / 60) : ($iHue / 60)

				Switch $nHuePre
					Case (-1) To 1.0
						$iRed = $nMaxRGB
						If $nHuePre < 0 Then
							$iGreen = $nMinRGB
							$iBlue = ($iGreen - $nHuePre * $nChroma)

						Else
							$iBlue = $nMinRGB
							$iGreen = ($iBlue + $nHuePre * $nChroma)
						EndIf

					Case 1.1 To 3.0
						$iGreen = $nMaxRGB
						If (($nHuePre - 2) < 0) Then
							$iBlue = $nMinRGB
							$iRed = ($iBlue - ($nHuePre - 2) * $nChroma)

						Else
							$iRed = $nMinRGB
							$iBlue = ($iRed + ($nHuePre - 2) * $nChroma)
						EndIf

					Case 3.1 To 5
						$iBlue = $nMaxRGB
						If (($nHuePre - 4) < 0) Then
							$iRed = $nMinRGB
							$iGreen = ($iRed - ($nHuePre - 4) * $nChroma)

						Else
							$iGreen = $nMinRGB
							$iRed = ($iGreen + ($nHuePre - 4) * $nChroma)
						EndIf
				EndSwitch

				$iRed = Round(($iRed * 255))
				$iGreen = Round(($iGreen * 255))
				$iBlue = Round(($iBlue * 255))

				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

				Return SetError($__LO_STATUS_SUCCESS, 3, $iLong) ; Return Long from HSB

			Else

				Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0) ; Wrong parameters
			EndIf

		Case 4 ; CMYK
			If Not (IsInt($vVal1) And IsInt($vVal2) And IsInt($vVal3) And IsInt($vVal4)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0) ; CMYK not integers.

			; CMYK to RGB
			$nCyan = ($vVal1 / 100)
			$nMagenta = ($vVal2 / 100)
			$nYellow = ($vVal3 / 100)
			$nBlack = ($vVal4 / 100)

			$iRed = Round((255 * (1 - $nBlack) * (1 - $nCyan)))
			$iGreen = Round((255 * (1 - $nBlack) * (1 - $nMagenta)))
			$iBlue = Round((255 * (1 - $nBlack) * (1 - $nYellow)))

			$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

			Return SetError($__LO_STATUS_SUCCESS, 4, $iLong) ; Long from CMYK

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0) ; wrong number of Parameters
	EndSwitch
EndFunc   ;==>_LO_ConvertColorToLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_DocConnect
; Description ...: Connect to an already opened instance of LibreOffice.
; Syntax ........: _LO_DocConnect([$iMode = $LO_DOC_CONNECT_MODE_CURRENT[, $sSearch = ""[, $bCaseless = False]]])
; Parameters ....: $iMode               - [optional] (0-4) Default is $LO_DOC_CONNECT_MODE_CURRENT. The Connect mode. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  $sSearch             - [optional] Default is "". The Name, Title or Path of the Document to search for. See remarks.
;                  $bCaseless           - [optional] Default is False. If True, searches are caseless when using $LO_DOC_CONNECT_MODE_SEARCH_* flags.
; Return values .: Success: Object or Array.
;                  @Error: 0, @Extended: ?, Return: Object = Success, The Object for the current, or last active document is returned. @Extended set to Document type Constant as an Integer. See Constants, $LO_DOC_TYPE_* as defined in LibreOffice_Constants.au3.
;                  @Error: 0, @Extended: ?, Return: Object = Success, The Object for the found Document with matching Name, Title or Path. @Extended set to Document type Constant as an Integer. See Constants, $LO_DOC_TYPE_* as defined in LibreOffice_Constants.au3.
;                  @Error: 0, @Extended: ?, Return: Array = Success, An Array of all open LibreOffice Documents. @Extended is set to number of results. See remarks.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $iMode not an Integer, less than 0 or greater than 4. See Constants, $LO_DOC_CONNECT_MODE_* as defined in LibreOffice_Constants.au3.
;                  @Error: 1, @Extended: 2 = $sSearch not a String.
;                  @Error: 1, @Extended: 3 = $bCaseless not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating ServiceManager object.
;                  @Error: 2, @Extended: 2 = Error creating Desktop object.
;                  @Error: 2, @Extended: 3 = Error creating enumeration of open documents.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = No open LibreOffice documents.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Document Object.
;                  @Error: 3, @Extended: 3 = Failed to identify Document type.
;                  @Error: 3, @Extended: 4 = Error converting path to LibreOffice URL.
;                  @Error: 3, @Extended: 5 = No matches found.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The value used for $sSearch depends on the flag called in $iMode. It is ignored except for the $LO_DOC_CONNECT_MODE_SEARCH_* flags.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_TITLE, $sSearch must be the full Title with Office and Component name; e.g: "Test.odt — LibreOffice Writer". This will be the same Title AutoIt would match or return from functions like WinGetTitle.
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME, $sSearch must be the Document's full name, without the extension; e.g: "Test".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT, $sSearch must be the Document's name, with the extension; e.g: "Test.odt". If the Document hasn't been saved, just the name will work, e.g., "Untitled 1".
;                  If $iMode is called with $LO_DOC_CONNECT_MODE_SEARCH_PATH, $sSearch must be the full Path of the document (Name and extension included); e.g: "C:\file\Test.odt."
;                  The Connect All option returns an array with two columns per result. ($aArray[0][2]), each result is stored in a separate row.
;                  -Row 1, Column 0 contains the Object for that document. e.g. $aArray[0][0] = $oDoc
;                  -Row 1, Column 1 contains the Document's Type as an Integer. See Constants, $LO_DOC_TYPE_* as defined in LibreOffice_Constants.au3. e.g. $aArray[0][1] = $LO_DOC_TYPE_CALC
;                  -Row 2, Column 0 contains the Object for the next document. e.g. $aArray[1][0] = $oDoc2. And so on.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_DocConnect($iMode = $LO_DOC_CONNECT_MODE_CURRENT, $sSearch = "", $bCaseless = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0, $iDocType
	Local $aoConnectAll[0][2]
	Local $sCaseless = ""
	Local $oEnumDoc, $oDoc, $oServiceManager, $oDesktop

	If Not __LO_IntIsBetween($iMode, $LO_DOC_CONNECT_MODE_ALL, $LO_DOC_CONNECT_MODE_SEARCH_PATH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sSearch) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCaseless) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	If Not $oDesktop.getComponents.hasElements() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; no L.O open

	Switch $iMode
		Case $LO_DOC_CONNECT_MODE_ALL
			$oEnumDoc = $oDesktop.getComponents.createEnumeration()
			If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				If (UBound($aoConnectAll) <= $iCount) Then ReDim $aoConnectAll[$iCount + 1][2]
				$aoConnectAll[$iCount][0] = $oDoc
				$aoConnectAll[$iCount][1] = _LO_DocGetType($oDoc)
				$iCount += 1
				Sleep((IsInt($iCount / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
			WEnd

			Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoConnectAll)

		Case $LO_DOC_CONNECT_MODE_CURRENT
			$oDoc = $oDesktop.currentComponent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$iDocType = _LO_DocGetType($oDoc)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.

			Return SetError($__LO_STATUS_SUCCESS, $iDocType, $oDoc)

		Case $LO_DOC_CONNECT_MODE_SEARCH_TITLE, $LO_DOC_CONNECT_MODE_SEARCH_NAME, $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT, $LO_DOC_CONNECT_MODE_SEARCH_PATH
			$sSearch = StringRegExpReplace($sSearch, "(^\s*|\s*$)", "")     ; Strip leading and trailing spaces

			If $bCaseless Then $sCaseless = "(?i)"

			If ($iMode = $LO_DOC_CONNECT_MODE_SEARCH_PATH) Then
				$sSearch = _LO_PathConvert($sSearch, $LO_PATHCONV_OFFICE_RETURN) ; Convert to L.O File path.
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
			EndIf

			$oEnumDoc = $oDesktop.getComponents.createEnumeration()
			If Not IsObj($oEnumDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			While $oEnumDoc.hasMoreElements()
				$oDoc = $oEnumDoc.nextElement()
				If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				Switch $iMode
					Case $LO_DOC_CONNECT_MODE_SEARCH_TITLE
						; First make sure Current Controller is available (It wont be if Document is opened Hidden, in some Components.).
						If IsObj($oDoc.CurrentController()) And StringRegExp($oDoc.CurrentController.Frame.Title(), $sCaseless & "\Q" & $sSearch & "\E") Then
							$iDocType = _LO_DocGetType($oDoc)
							If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.

							Return SetError($__LO_STATUS_SUCCESS, $iDocType, $oDoc)
						EndIf

					Case $LO_DOC_CONNECT_MODE_SEARCH_NAME
						; Add spaces after name in case user put some in the Document name.
						; Add additional capture for Extension to just match the name the user put in, else force match at end of String for unsaved Documents.
						If StringRegExp($oDoc.Title(), $sCaseless & "\Q" & $sSearch & "\E\s*(\.\w+)?$") Then
							$iDocType = _LO_DocGetType($oDoc)
							If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.

							Return SetError($__LO_STATUS_SUCCESS, $iDocType, $oDoc)
						EndIf

					Case $LO_DOC_CONNECT_MODE_SEARCH_NAME_WITH_EXT
						If StringRegExp($oDoc.Title(), $sCaseless & "\Q" & $sSearch & "\E") Then
							$iDocType = _LO_DocGetType($oDoc)
							If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.

							Return SetError($__LO_STATUS_SUCCESS, $iDocType, $oDoc)
						EndIf

					Case $LO_DOC_CONNECT_MODE_SEARCH_PATH
						If StringRegExp($oDoc.getURL(), $sCaseless & "\Q" & $sSearch & "\E") Then
							$iDocType = _LO_DocGetType($oDoc)
							If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Failed to identify Doc type.

							Return SetError($__LO_STATUS_SUCCESS, $iDocType, $oDoc)
						EndIf
				EndSwitch
			WEnd
	EndSwitch

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; No matches
EndFunc   ;==>_LO_DocConnect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_DocGetType
; Description ...: Identify the document's type.
; Syntax ........: _LO_DocGetType(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous Document Open, Connect, or Create function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning the document's type as an Integer. See Constants, $LO_DOC_TYPE_* as defined in LibreOffice_Constants.au3.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve a Form Cursor.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Table or Query name.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Active Connection Object.
;                  @Error: 3, @Extended: 4 = Failed to retrieve Document Creation Arguments Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_DocGetType(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iDocType = $LO_DOC_TYPE_UNKNOWN
	Local $atArgs
	Local $asMatch
	Local $oRowSet, $oActiveConn
	Local $sName, $sDocumentBaseURL, $sDocTitle = ""
	Local Const $__STR_REGEXPARRAYMATCH = 1
	Local Const $sBaseServiceName = "com.sun.star.sdb.OfficeDatabaseDocument", _ ; Base
			$sBaseQueryDesignServiceName = "com.sun.star.sdb.QueryDesign", _ ; Base Query Document in Design mode.
			$sBaseTableDesignServiceName = "com.sun.star.sdb.TableDesign", _ ; Base Table Document in Design mode.
			$sBaseReportDesignServiceName = "com.sun.star.report.ReportDefinition", _ ; Base Report Doc in Design mode.
			$sBaseSubServiceName = "com.sun.star.sdb.DataSourceBrowser", _ ; Could be a Query or a Table Document in Viewing mode.
			$sBasicIDEServiceName = "com.sun.star.script.BasicIDE", _ ; Basic IDE
			$sCalcServiceName = "com.sun.star.sheet.SpreadsheetDocument", _ ; Calc or Report (View)
			$sDrawServiceName = "com.sun.star.drawing.DrawingDocument", _ ; Draw
			$sImpressServiceName = "com.sun.star.presentation.PresentationDocument", _ ; Impress
			$sMathServiceName = "com.sun.star.formula.FormulaProperties", _ ; Math
			$sWriterWebServiceName = "com.sun.star.text.WebDocument", _ ; Writer Web/HTML
			$sTextDocServiceName = "com.sun.star.text.TextDocument" ; Could be a Writer Doc, or Form (View and Design), or Report (View)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If $oDoc.supportsService($sBaseServiceName) Then
		$iDocType = $LO_DOC_TYPE_BASE

	ElseIf $oDoc.supportsService($sBaseQueryDesignServiceName) Then
		$iDocType = $LO_DOC_TYPE_BASE_QUERY_DESIGN

	ElseIf $oDoc.supportsService($sBaseTableDesignServiceName) Then
		$iDocType = $LO_DOC_TYPE_BASE_TABLE_DESIGN

	ElseIf $oDoc.supportsService($sBaseReportDesignServiceName) Then
		$iDocType = $LO_DOC_TYPE_BASE_REPORT_DESIGN

	ElseIf $oDoc.supportsService($sBaseSubServiceName) Then
		$oRowSet = $oDoc.FormOperations.Cursor()
		If Not IsObj($oRowSet) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$sName = $oRowSet.Command()
		If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oActiveConn = $oRowSet.ActiveConnection()
		If Not IsObj($oActiveConn) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		If $oActiveConn.Queries.hasByName($sName) Then
			$iDocType = $LO_DOC_TYPE_BASE_QUERY_VIEW

		ElseIf $oActiveConn.Tables.hasByName($sName) Then
			$iDocType = $LO_DOC_TYPE_BASE_TABLE_VIEW
		EndIf

	ElseIf $oDoc.supportsService($sTextDocServiceName) Then
		; For Form Doc, check if it has a parent, then check if parent is a Base Doc.
		; Form documents have DocumentBaseURL also, URL matches Base file URL.
		; If Form is read-Only, then View mode, else Design.

		If IsObj($oDoc.Parent()) And $oDoc.Parent.supportsService($sBaseServiceName) Then         ; A Form, View or Design.
			$atArgs = $oDoc.Args()
			If Not IsArray($atArgs) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			For $i = 0 To UBound($atArgs) - 1
				If ($atArgs[$i].Name() = "DocumentBaseURL") Then
					$sDocumentBaseURL = $atArgs[$i].Value()
					ExitLoop
				EndIf
				Sleep((IsInt($i / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
			Next

			If IsString($sDocumentBaseURL) And ($oDoc.Parent.URL() = $sDocumentBaseURL) Then
				If $oDoc.isReadOnly() Then         ; Form in View mode.
					$iDocType = $LO_DOC_TYPE_BASE_FORM_VIEW

				Else         ; Form in Design mode.
					$iDocType = $LO_DOC_TYPE_BASE_FORM_DESIGN
				EndIf
			EndIf

		ElseIf $oDoc.isReadOnly() Then         ; Could be a Writer Doc or Report in View mode, as both have no parent.
			$atArgs = $oDoc.Args()
			If Not IsArray($atArgs) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			For $i = 0 To UBound($atArgs) - 1
				If ($atArgs[$i].Name() = "DocumentBaseURL") Then
					$sDocumentBaseURL = $atArgs[$i].Value()
					ExitLoop
				EndIf
				Sleep((IsInt($i / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
			Next

			If IsString($sDocumentBaseURL) Then         ; Probably a Report in View mode.
				; Title: "QryAutoIt.docx (read-only)" (Need to cut it at the space.)
				$asMatch = StringRegExp($oDoc.Title(), "^(.+\.\w+)", $__STR_REGEXPARRAYMATCH)

				If Not @error Then $sDocTitle = $asMatch[0]

				; Report doc in view mode has URL: file:///C:/Users/Owner/AppData/Local/Temp/lu1376bt3zd.tmp/QryAutoIt.docx
				; Can search for .tmp/ + Report name.
				If StringRegExp($sDocumentBaseURL, "\/\w+\.(?i)tmp\/\Q" & $sDocTitle & "\E") Then $iDocType = $LO_DOC_TYPE_BASE_REPORT_VIEW

			Else         ; Assuming, and most probably, a Writer Document.
				$iDocType = $LO_DOC_TYPE_WRITER
			EndIf

		Else         ; Not Read Only, most likely a Writer Doc.
			$iDocType = $LO_DOC_TYPE_WRITER
		EndIf

	ElseIf $oDoc.supportsService($sBasicIDEServiceName) Then
		$iDocType = $LO_DOC_TYPE_BASIC_IDE

	ElseIf $oDoc.supportsService($sCalcServiceName) Then
		If $oDoc.isReadOnly() Then         ; Could be a Calc Doc or Report in View mode.
			$atArgs = $oDoc.Args()
			If Not IsArray($atArgs) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			For $i = 0 To UBound($atArgs) - 1
				If ($atArgs[$i].Name() = "DocumentBaseURL") Then
					$sDocumentBaseURL = $atArgs[$i].Value()
					ExitLoop
				EndIf
				Sleep((IsInt($i / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
			Next

			If IsString($sDocumentBaseURL) Then         ; Probably a Report in View mode.
				; Title: "QryAutoIt1.ods (read-only)" (Need to cut it at the space.)
				$asMatch = StringRegExp($oDoc.Title(), "^(.+\.\w+)", $__STR_REGEXPARRAYMATCH)

				If Not @error Then $sDocTitle = $asMatch[0]

				; Report doc in view mode has URL: file:///C:/Users/Owner/AppData/Local/Temp/lu20921b6g.tmp/QryAutoIt1.ods
				; Can search for .tmp/ + Report name.
				If StringRegExp($sDocumentBaseURL, "\/\w+\.(?i)tmp\/\Q" & $sDocTitle & "\E") Then $iDocType = $LO_DOC_TYPE_BASE_REPORT_VIEW

			Else         ; Assuming, and most probably, a Calc Document.
				$iDocType = $LO_DOC_TYPE_CALC
			EndIf

		Else         ; Not Read Only, most probably a Calc Document.
			$iDocType = $LO_DOC_TYPE_CALC
		EndIf

	ElseIf $oDoc.supportsService($sImpressServiceName) Then         ; ALWAYS need to check for Impress before Draw, as Draw has all same services as Impress.
		$iDocType = $LO_DOC_TYPE_IMPRESS

	ElseIf $oDoc.supportsService($sDrawServiceName) Then
		$iDocType = $LO_DOC_TYPE_DRAW

	ElseIf $oDoc.supportsService($sMathServiceName) Then
		$iDocType = $LO_DOC_TYPE_MATH

	ElseIf $oDoc.supportsService($sWriterWebServiceName) Then
		$iDocType = $LO_DOC_TYPE_WRITER_WEB
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $iDocType)
EndFunc   ;==>_LO_DocGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_GradientMulticolorAdd
; Description ...: Add a ColorStop to a Gradient ColorStop Array.
; Syntax ........: _LO_GradientMulticolorAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
; Parameters ....: $avColorStops        - A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - The array index to insert the color stop. 0 Based. Call the last element index plus 1 to insert at the end.
;                  $nStopOffset         - (0-1.0) The ColorStop offset value.
;                  $iColor              - (0-16777215) The ColorStop color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. ColorStop successfully added to array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $avColorStops not an Array.
;                  @Error: 1, @Extended: 2 = $avColorStops does not contain two columns.
;                  @Error: 1, @Extended: 3 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
;                  @Error: 1, @Extended: 4 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error: 1, @Extended: 5 = $iColor not an Integer, less than 0 or greater than 16777215.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_GradientMulticolorAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	ReDim $avColorStops[UBound($avColorStops) + 1][2]

	For $iToWrite = (UBound($avColorStops) - 1) To 0 Step -1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iColor
			ExitLoop

		Else
			$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite - 1][0]
			$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite - 1][1]
		EndIf

		Sleep((IsInt($iToWrite / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_GradientMulticolorAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_GradientMulticolorDelete
; Description ...: Delete a ColorStop from a Gradient ColorStop Array.
; Syntax ........: _LO_GradientMulticolorDelete(ByRef $avColorStops, $iIndex)
; Parameters ....: $avColorStops        - A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - The array index to delete. 0 Based.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. ColorStop successfully removed from array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $avColorStops not an Array.
;                  @Error: 1, @Extended: 2 = $avColorStops does not contain two columns.
;                  @Error: 1, @Extended: 3 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_GradientMulticolorDelete(ByRef $avColorStops, $iIndex)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2
	Local $iToRead = 0

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 2
		If $iToWrite = $iIndex Then $iToRead += 1

		$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite + $iToRead][0]
		$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite + $iToRead][1]

		Sleep((IsInt($iToWrite / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	ReDim $avColorStops[UBound($avColorStops) - 1][2]

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_GradientMulticolorDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_GradientMulticolorModify
; Description ...: Modify a ColorStop in a Gradient ColorStop Array.
; Syntax ........: _LO_GradientMulticolorModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
; Parameters ....: $avColorStops        - A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - The array index to modify. 0 Based.
;                  $nStopOffset         - (0-1.0) The ColorStop offset value.
;                  $iColor              - (0-16777215) The ColorStop color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. ColorStop successfully modified.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $avColorStops not an Array.
;                  @Error: 1, @Extended: 2 = $avColorStops does not contain two columns.
;                  @Error: 1, @Extended: 3 = $iIndex not an Integer, less than 0 or greater than last element.
;                  @Error: 1, @Extended: 4 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error: 1, @Extended: 5 = $iColor not an Integer, less than 0 or greater than 16777215.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_GradientMulticolorModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iColor
			ExitLoop
		EndIf

		Sleep((IsInt($iToWrite / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_GradientMulticolorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_InitializePortable
; Description ...: Setup Portable LibreOffice (Or Open Office) for use in this UDF.
; Syntax ........: _LO_InitializePortable($sOfficePortablePath)
; Parameters ....: $sOfficePortablePath - The Path to the Portable LibreOffice/OpenOffice folder. See remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Portable LibreOffice/OpenOffice ServiceManager successfully created and stored.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $sOfficePortablePath not a String.
;                  @Error: 1, @Extended: 2 = Folder called in $sOfficePortablePath does not contain the App folder. Perhaps wrong directory?
;                  @Error: 1, @Extended: 3 = soffice.exe not found in $sOfficePortablePath\App\libreoffice\program\ or $sOfficePortablePath\App\openoffice\program\.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to clear stored Portable LO/OO ServiceManager.
;                  @Error: 3, @Extended: 2 = Failed to initialize portable LibreOffice ServiceManager.
;                  @Error: 3, @Extended: 3 = Failed to initialize portable OpenOffice ServiceManager.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The path called in $sOfficePortablePath should be to the Portable LibreOffice folder containing the shortcuts to each element, and also the "App", "Data" and "Other" folders. e.g. C:\LibreOfficePortablePrevious
;                  This UDF hasn't been thoroughly tested using portable LibreOffice, but rather the Installed version. Make sure to test all things yourself!
;                  So far, the following method is the only way I've been able to successfully initialize the portable LibreOffice version from AutoIt. How it works is as follows:
;                  1a. If LibreOffice/OpenOffice is already installed on the system, an instance of com.sun.star.ServiceManager is started. __OR__
;                  1b. If LibreOffice/OpenOffice is NOT already installed on the system, a handful of temporary registry entries are created in HKEY_CURRENT_USER, to allow an instance of com.sun.star.ServiceManager to be started.
;                  2. The Portable LibreOffice/OpenOffice (LO/OO) is started in --headless mode as a listening server, using the flag --accept.
;                  3. The ServiceManager created in step 1 is used to create a UNO URL resolver, which is used to obtain a ServiceManager Object from the listening Portable LO/OO.
;                  4. The retrieved Portable LO/OO ServiceManager is stored as a static variable for future use. The Portable LO/OO path is also stored as a static variable in case I need to re-create the ServiceManager.
;                  5a. The ServiceManager created from the registry is no longer used.
;                  5b. The ServiceManager created from the registry is no longer used, and the temporary Registry entries created in HKEY_CURRENT_USER are (hopefully) deleted.
;                  If the COM error "Binary URP bridge already disposed" is encountered, all instances of soffice.exe or soffice.bin must be closed with TaskManager.
;                  If running this with an installed version of LibreOffice present the flag SingleAppInstance may need to be set to False in the "LibreOfficePortablePrevious.ini" [or similar name], found at: C:\LibreOfficePortablePrevious\App\AppInfo\Launcher\LibreOfficePortablePrevious.ini.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LO_InitializePortable($sOfficePortablePath)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsString($sOfficePortablePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$sOfficePortablePath = StringRegExpReplace($sOfficePortablePath, "^\s+|\s+$", "") ; Strip beginning and ending spaces.
	If ($sOfficePortablePath <> "") And Not FileExists($sOfficePortablePath & "\App") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Make sure we're starting from the right folder.

	If ($sOfficePortablePath = "") Then
		__LO_SetPortableServiceManager($sOfficePortablePath)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ElseIf FileExists($sOfficePortablePath & "\App\libreoffice\program\soffice.exe") Then ; Check LibreOffice path.
		__LO_SetPortableServiceManager($sOfficePortablePath & "\App\libreoffice\program\soffice.exe")
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	ElseIf FileExists($sOfficePortablePath & "\App\openoffice\program\soffice.exe") Then ; Check OpenOffice path.
		__LO_SetPortableServiceManager($sOfficePortablePath & "\App\openoffice\program\soffice.exe")
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_InitializePortable

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_PathConvert
; Description ...: Converts the input path to or from a LibreOffice URL notation path.
; Syntax ........: _LO_PathConvert($sFilePath[, $iReturnMode = $LO_PATHCONV_AUTO_RETURN])
; Parameters ....: $sFilePath           - Full path to convert in String format.
;                  $iReturnMode         - [optional] (0-2) Default is $__g_iAutoReturn. The type of path format to return. See Constants, $LO_PATHCONV_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: String.
;                  @Error: 0, @Extended: 1, Return: String = Returning converted File Path from LibreOffice URL.
;                  @Error: 0, @Extended: 2, Return: String = Returning converted path from File Path to LibreOffice URL.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $sFilePath is not a string
;                  @Error: 1, @Extended: 2 = $iReturnMode not a Integer, less than 0 or greater than 2. See constants, $LO_PATHCONV_* as defined in LibreOffice_Constants.au3..
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice URL notation is based on the Internet Standard RFC 1738, which means only [0-9],[a-zA-Z] are allowed in paths, most other characters need to be converted into ISO 8859-1 (ISO Latin) such as is found in internet URL's (spaces become %20).
;                  See: StarOfficeTM 6.0 Office SuiteA SunTM ONE Software Offering, Basic Programmer's Guide; Page 74
;                  The user generally should not even need this function, as I have endeavored to convert any URLs to the appropriate computer path format and any input computer paths to a LibreOffice URL.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_PathConvert($sFilePath, $iReturnMode = $LO_PATHCONV_AUTO_RETURN)
	Local Const $__STR_STRIPLEADING = 1
	Local $asURLReplace[9][2] = [["%", "%25"], [" ", "%20"], ["\", "/"], [";", "%3B"], ["#", "%23"], ["^", "%5E"], ["{", "%7B"], ["}", "%7D"], ["`", "%60"]]
	Local $iPathSearch, $iFileSearch, $iPartialPCPath, $iPartialFilePath

	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iReturnMode, $LO_PATHCONV_AUTO_RETURN, $LO_PATHCONV_PCPATH_RETURN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sFilePath = StringStripWS($sFilePath, $__STR_STRIPLEADING)

	$iPathSearch = StringRegExp($sFilePath, "(?i)\b[A-Z]:\\") ; Search For a Computer Path, as in C:\ etc.
	$iPartialPCPath = StringInStr($sFilePath, "\") ; Search for partial computer Path containing a backslash.
	$iFileSearch = StringInStr($sFilePath, "file:///", 0, 1, 1, 9) ; Search for a full LibreOffice path, which begins with File:///
	$iPartialFilePath = StringInStr($sFilePath, "/") ; Search For a Partial LibreOffice path containing forward slash

	If ($iReturnMode = $LO_PATHCONV_AUTO_RETURN) Then
		If ($iPathSearch > 0) Or ($iPartialPCPath > 0) Then ;  if file path contains partial or full PC path, set to convert to LibreOffice URL.
			$iReturnMode = $LO_PATHCONV_OFFICE_RETURN

		ElseIf ($iFileSearch > 0) Or ($iPartialFilePath > 0) Then ;  if file path contains partial or full LibreOffice URL, set to convert to PC Path.
			$iReturnMode = $LO_PATHCONV_PCPATH_RETURN

		Else ; If file path contains neither above. convert to LibreOffice URL
			$iReturnMode = $LO_PATHCONV_OFFICE_RETURN
		EndIf
	EndIf

	Switch $iReturnMode
		Case $LO_PATHCONV_OFFICE_RETURN
			If $iFileSearch > 0 Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)
			If ($iPathSearch > 0) Then $sFilePath = "file:///" & $sFilePath

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][0], $asURLReplace[$i][1])
				Sleep((IsInt($i / $__LOCONST_SLEEP_DIV)) ? (10) : (0))
			Next

			Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)

		Case $LO_PATHCONV_PCPATH_RETURN
			If ($iPathSearch > 0) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)
			If ($iFileSearch > 0) Then $sFilePath = StringReplace($sFilePath, "file:///", Null)

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][1], $asURLReplace[$i][0])
				Sleep((IsInt($i / $__LOCONST_SLEEP_DIV)) ? (10) : (0))
			Next

			Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)
	EndSwitch
EndFunc   ;==>_LO_PathConvert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_PrintersGetNames
; Description ...: Enumerates all installed printers, or current default printer.
; Syntax ........: _LO_PrintersGetNames([$bDefaultOnly = False])
; Parameters ....: $bDefaultOnly        - [optional] Default is False. If True, returns only the name of the current default printer. LibreOffice 6.3 and up only.
; Return values .: Success: An array or String.
;                  @Error: 0, @Extended: 1, Return: String = Returning the default printer's name.
;                  @Error: 0, @Extended: ?, Return: Array = Returning an array of strings of all installed printers' names. @Extended set to number of results.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $bDefaultOnly not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failure Creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Failure creating "com.sun.star.awt.PrinterServer" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Default printer name.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Array of printer names.
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 4.1.
;                  @Error: 6, @Extended: 2 = Current LibreOffice version lower than 6.3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function works for LibreOffice 4.1 and Up.
; Related .......: _LO_PrintersGetNamesAlt
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_PrintersGetNames($bDefaultOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oPrintServer
	Local $sDefault
	Local $asPrinters[0]

	If Not __LO_VersionCheck(4.1) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsBool($bDefaultOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oServiceManager = __LO_ServiceManager()
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oPrintServer = $oServiceManager.createInstance("com.sun.star.awt.PrinterServer")
	If Not IsObj($oPrintServer) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If $bDefaultOnly Then
		If Not __LO_VersionCheck(6.3) Then Return SetError($__LO_STATUS_VER_ERROR, 2, 0)

		$sDefault = $oPrintServer.getDefaultPrinterName()
		If IsString($sDefault) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sDefault)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	$asPrinters = $oPrintServer.getPrinterNames()
	If Not IsArray($asPrinters) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asPrinters), $asPrinters)
EndFunc   ;==>_LO_PrintersGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_PrintersGetNamesAlt
; Description ...: Alternate function; Enumerates all installed printers, or current default printer.
; Syntax ........: _LO_PrintersGetNamesAlt([$sPrinterName = ""[, $bReturnDefault = False]])
; Parameters ....: $sPrinterName        - [optional] Default is "". Name of the printer to list. Default "" returns the list of all printers. See Remarks.
;                  $bReturnDefault      - [optional] Default is False. If True, returns only the name of the current default printer.
; Return values .: Success: Array or String.
;                  @Error: 0, @Extended: 1, Return: String = Returning the default printer name. See remarks. @Extended is set to the number of results.
;                  @Error: 0, @Extended: ?, Return: Array = Returning an array of strings containing all installed printers. See remarks. Number of results returned in @Extended.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $sPrinterName not a String.
;                  @Error: 1, @Extended: 2 = $bReturnDefault not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failure Creating Object.
;                  @Error: 2, @Extended: 2 = Failure retrieving printer list Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve default printer name.
;                  --Printer Related Errors--
;                  @Error: 5, @Extended: 1 = No default printer found.
; Author ........: jguinch (_PrintMgr_EnumPrinter)
; Modified ......: donnyh13 - Added input error checking. Added a return default printer only option.
; Remarks .......: When $bReturnDefault is False, The function returns all installed printers for the user running the script in an array.
;                  If $sPrinterName is set, the name must be exact, or no results will be found, unless you use an asterisk (*) for partial name searches, either prefixed (*Canon), suffixed (Canon*), or both (*Canon*).
;                  When $bReturnDefault is True, The function returns only the default printer's name or sets an error if no default printer is found.
; Related .......: _LO_PrintersGetNames
; Link ..........: (Printmgr.au3) https://www.autoitscript.com/forum/topic/155485-printers-management-udf/
; Example .......: Yes
; ===============================================================================================================================
Func _LO_PrintersGetNamesAlt($sPrinterName = "", $bReturnDefault = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asPrinterNames[10]
	Local $sFilter
	Local $iCount = 0
	Local $sName
	Local Const $wbemFlagReturnImmediately = 0x10, $wbemFlagForwardOnly = 0x20
	Local $oWMIService, $oPrinters

	If Not IsString($sPrinterName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnDefault) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $sPrinterName <> "" Then $sFilter = StringReplace(" Where Name like '" & StringReplace($sPrinterName, "\", "\\") & "'", "*", "%")
	$oWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	If Not IsObj($oWMIService) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oPrinters = $oWMIService.ExecQuery("Select * from Win32_Printer" & $sFilter, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If Not IsObj($oPrinters) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	For $oPrinter In $oPrinters
		If $bReturnDefault Then
			If $oPrinter.Default() Then
				$sName = $oPrinter.Name()
				If Not IsString($sName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				Return SetError($__LO_STATUS_SUCCESS, 1, $sName)
			EndIf

		Else
			If $iCount >= (UBound($asPrinterNames) - 1) Then ReDim $asPrinterNames[UBound($asPrinterNames) * 2]
			$asPrinterNames[$iCount] = $oPrinter.Name()
			$iCount += 1
		EndIf
	Next
	If $bReturnDefault Then Return SetError($__LO_STATUS_PRINTER_RELATED_ERROR, 1, 0)

	ReDim $asPrinterNames[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asPrinterNames)
EndFunc   ;==>_LO_PrintersGetNamesAlt

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_Terminate
; Description ...: Closes the background instance of LibreOffice.
; Syntax ........: _LO_Terminate([$bForceClose = False[, $iSleep = 250]])
; Parameters ....: $bForceClose         - [optional] Default is False. If True, any opened documents will be closed. See remarks.
;                  $iSleep              - [optional] Default is 250. The amount of time to sleep before perofrming the terminate command, in milliseconds. See remarks.
; Return values .: Success: Boolean
;                  @Error: 0, @Extended: 0, Return: Boolean = Success. Terminate command was successfuly processed. Returning True if all Documents agree to be terminated.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $bForceClose not a Boolean.
;                  @Error: 1, @Extended: 2 = $iSleep not an Integer or less than 0.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a ServiceManager Object.
;                  @Error: 2, @Extended: 2 = Failed to create a com.sun.star.frame.Desktop Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If $bForceClose is called with False, and there are no open Documents, the background instance of soffice.bin will be terminated.
;                  If $bForceClose is called with True, all opened documents are closed, any documents with unsaved changes will have a save dialog initiated for the user to interact with.
;                  If this function was not used, a left-over instance of soffice.bin would remain running after automating LibreOffice.
;                  Some Online sources recommend to allow a minimum of 500ms sleep before terminating the LibreOffice instance to allow it finish closing any documents etc., otherwise the "Document Recovery" mode will be triggered upon next startup.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_Terminate($bForceClose = False, $iSleep = 250)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $oDesktop
	Local $bTerminated

	If Not IsBool($bForceClose) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iSleep, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not $oDesktop.getComponents.hasElements() Or $bForceClose Then ; no L.O open, or force it to close.
		Sleep($iSleep) ; Sleep to make sure LO has time to finish any processes it may be doing, otherwise document recovery mode may be triggered at the next startup.
		$bTerminated = $oDesktop.Terminate()
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $bTerminated)
EndFunc   ;==>_LO_Terminate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_TransparencyGradientMultiAdd
; Description ...: Add a ColorStop to a Gradient ColorStop Array.
; Syntax ........: _LO_TransparencyGradientMultiAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
; Parameters ....: $avColorStops        - A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - The array index to insert the color stop. 0 Based. Call the last element index plus 1 to insert at the end.
;                  $nStopOffset         - (0-1.0) The ColorStop offset value.
;                  $iTransparency       - (0-100) The ColorStop Transparency value percentage. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. ColorStop successfully added to array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $avColorStops not an Array.
;                  @Error: 1, @Extended: 2 = $avColorStops does not contain two columns.
;                  @Error: 1, @Extended: 3 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
;                  @Error: 1, @Extended: 4 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error: 1, @Extended: 5 = $iTransparency not an Integer, less than 0 or greater than 100.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_TransparencyGradientMultiAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	ReDim $avColorStops[UBound($avColorStops) + 1][2]

	For $iToWrite = (UBound($avColorStops) - 1) To 0 Step -1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iTransparency
			ExitLoop

		Else
			$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite - 1][0]
			$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite - 1][1]
		EndIf

		Sleep((IsInt($iToWrite / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_TransparencyGradientMultiAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_TransparencyGradientMultiDelete
; Description ...: Delete a ColorStop from a Gradient ColorStop Array.
; Syntax ........: _LO_TransparencyGradientMultiDelete(ByRef $avColorStops, $iIndex)
; Parameters ....: $avColorStops        - A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - The array index to delete. 0 Based.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. ColorStop successfully removed from array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $avColorStops not an Array.
;                  @Error: 1, @Extended: 2 = $avColorStops does not contain two columns.
;                  @Error: 1, @Extended: 3 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_TransparencyGradientMultiDelete(ByRef $avColorStops, $iIndex)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2
	Local $iToRead = 0

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 2
		If $iToWrite = $iIndex Then $iToRead += 1

		$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite + $iToRead][0]
		$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite + $iToRead][1]

		Sleep((IsInt($iToWrite / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	ReDim $avColorStops[UBound($avColorStops) - 1][2]

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_TransparencyGradientMultiDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_TransparencyGradientMultiModify
; Description ...: Modify a ColorStop in a Gradient ColorStop Array.
; Syntax ........: _LO_TransparencyGradientMultiModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
; Parameters ....: $avColorStops        - A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - The array index to modify. 0 Based.
;                  $nStopOffset         - (0-1.0) The ColorStop offset value.
;                  $iTransparency       - (0-100) The ColorStop Transparency value percentage. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. ColorStop successfully modified.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $avColorStops not an Array.
;                  @Error: 1, @Extended: 2 = $avColorStops does not contain two columns.
;                  @Error: 1, @Extended: 3 = $iIndex not an Integer, less than 0 or greater than last element.
;                  @Error: 1, @Extended: 4 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error: 1, @Extended: 5 = $iTransparency not an Integer, less than 0 or greater than 100.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_TransparencyGradientMultiModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iTransparency
			ExitLoop
		EndIf

		Sleep((IsInt($iToWrite / $__LOCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LO_TransparencyGradientMultiModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_UnitConvert
; Description ...: For converting measurement units.
; Syntax ........: _LO_UnitConvert($nValue, $iReturnType)
; Parameters ....: $nValue              - The Number to be converted.
;                  $iReturnType         - (0-10) The conversion type to perform on $nValue. See Constants, $LO_CONVERT_UNIT_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: Integer or Number.
;                  @Error: 0, @Extended: 1, Return: Number = Returning Number converted from TWIPS to Centimeters.
;                  @Error: 0, @Extended: 2, Return: Number = Returning Number converted from TWIPS to Inches.
;                  @Error: 0, @Extended: 3, Return: Integer = Returning Number converted from Millimeters to Hundredths of a Millimeter (HMM).
;                  @Error: 0, @Extended: 4, Return: Number = Returning Number converted from Hundredths of a Millimeter (HMM) to MM
;                  @Error: 0, @Extended: 5, Return: Integer = Returning Number converted from Centimeters To Hundredths of a Millimeter (HMM)
;                  @Error: 0, @Extended: 6, Return: Number = Returning Number converted from Hundredths of a Millimeter (HMM) To CM
;                  @Error: 0, @Extended: 7, Return: Integer = Returning Number converted from Inches to Hundredths of a Millimeter (HMM).
;                  @Error: 0, @Extended: 8, Return: Number = Returning Number converted from Hundredths of a Millimeter (HMM) to Inches.
;                  @Error: 0, @Extended: 9, Return: Integer = Returning Number converted from TWIPS to Hundredths of a Millimeter (HMM).
;                  @Error: 0, @Extended: 10, Return: Integer = Returning Number converted from Point to Hundredths of a Millimeter (HMM).
;                  @Error: 0, @Extended: 11, Return: Number = Returning Number converted from Hundredths of a Millimeter (HMM) to Point.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $nValue is not a Number.
;                  @Error: 1, @Extended: 2 = $iReturnType is not a Integer, less than 0 or greater than 10. See Constants, $LO_CONVERT_UNIT_* as defined in LibreOffice_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iReturnType does not match constants, See Constants, $LO_CONVERT_UNIT_* as defined in LibreOffice_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Hundredths of a Millimeter (HMM), is used in almost all LibreOffice functions that contain a measurement parameter.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LO_UnitConvert($nValue, $iReturnType)
	Local $iHMM, $iMM, $iCM, $iInch

	If Not IsNumber($nValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iReturnType, $LO_CONVERT_UNIT_TWIPS_CM, $LO_CONVERT_UNIT_HMM_PT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch $iReturnType
		Case $LO_CONVERT_UNIT_TWIPS_CM ; TWIPS TO CM
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			; 1 Inch = 2.54 CM
			$iCM = Round(Round($iInch * 2.54, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 1, Number($iCM))

		Case $LO_CONVERT_UNIT_TWIPS_INCH ; TWIPS to Inch
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			$iInch = Round(Round($iInch, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 2, Number($iInch))

		Case $LO_CONVERT_UNIT_MM_HMM ; Millimeter to Hundredths of a Millimeter (HMM).
			$iHMM = ($nValue * 100)
			$iHMM = Round(Round($iHMM, 1))

			Return SetError($__LO_STATUS_SUCCESS, 3, Number($iHMM))

		Case $LO_CONVERT_UNIT_HMM_MM ; Hundredths of a Millimeter (HMM) to Millimeter
			$iMM = ($nValue / 100)
			$iMM = Round(Round($iMM, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 4, Number($iMM))

		Case $LO_CONVERT_UNIT_CM_HMM ; Centimeter to Hundredths of a Millimeter (HMM)
			$iHMM = ($nValue * 1000)
			$iHMM = Round(Round($iHMM, 1))

			Return SetError($__LO_STATUS_SUCCESS, 5, Int($iHMM))

		Case $LO_CONVERT_UNIT_HMM_CM ; Hundredths of a Millimeter (HMM) to Centimeter
			$iCM = ($nValue / 1000)
			$iCM = Round(Round($iCM, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 6, Number($iCM))

		Case $LO_CONVERT_UNIT_INCH_HMM ; Inch to Hundredths of a Millimeter (HMM)
			; 1 Inch - 2.54 Cm; Hundredths of a Millimeter (HMM) = 1/1000 CM
			$iHMM = ($nValue * 2.54) * 1000
			$iHMM = Round(Round($iHMM, 1))

			Return SetError($__LO_STATUS_SUCCESS, 7, Int($iHMM))

		Case $LO_CONVERT_UNIT_HMM_INCH ; Hundredths of a Millimeter (HMM) to Inch
			; 1 Inch - 2.54 Cm; Hundredths of a Millimeter (HMM) = 1/1000 CM
			$iInch = ($nValue / 1000) / 2.54
			$iInch = Round(Round($iInch, 3), 2)

			Return SetError($__LO_STATUS_SUCCESS, 8, $iInch)

		Case $LO_CONVERT_UNIT_TWIPS_HMM ; TWIPS to Hundredths of a Millimeter (HMM)
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = (($nValue / 20) / 72)
			$iInch = Round(Round($iInch, 3), 2)
			; 1 Inch = 25.4 MM; 100 Hundredths of a Millimeter (HMM) = 1 MM
			$iHMM = Round($iInch * 25.4 * 100)

			Return SetError($__LO_STATUS_SUCCESS, 9, Int($iHMM))

		Case $LO_CONVERT_UNIT_PT_HMM
			; 1 pt = 35 Hundredths of a Millimeter (HMM)
			If ($nValue <> 0) Then $nValue = Round(($nValue * 35.2778))

			Return SetError($__LO_STATUS_SUCCESS, 10, $nValue)

		Case $LO_CONVERT_UNIT_HMM_PT
			If ($nValue <> 0) Then $nValue = Round(($nValue / 35.2778), 2)

			Return SetError($__LO_STATUS_SUCCESS, 11, $nValue)

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndSwitch
EndFunc   ;==>_LO_UnitConvert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LO_VersionGet
; Description ...: Retrieve the current Office version.
; Syntax ........: _LO_VersionGet([$bSimpleVersion = False[, $bReturnName = False]])
; Parameters ....: $bSimpleVersion      - [optional] Default is False. If True, returns a two digit version number, such as "7.3", else returns the complex version number, such as "7.3.2.4".
;                  $bReturnName         - [optional] Default is True. If True returns the Program Name, such as "LibreOffice", appended by the version, i.e. "LibreOffice 7.3".
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning the Office version in String format.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $bSimpleVersion not a Boolean.
;                  @Error: 1, @Extended: 2 = $bReturnName not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error: 2, @Extended: 2 = Error creating "com.sun.star.configuration.ConfigurationProvider" Object.
;                  @Error: 2, @Extended: 3 = Error creating property value.
; Author ........: Laurent Godard as found in Andrew Pitonyak's book; Zizi64 as found on OpenOffice forum.
; Modified ......: donnyh13, modified for AutoIt compatibility and error checking.
; Remarks .......: From Macro code by Zizi64 found at: https://forum.openoffice.org/en/forum/viewtopic.php?t=91542&sid=7f452d65e58ac1cd3cc6063350b5ada0
;                  And Andrew Pitonyak in "Useful Macro Information For OpenOffice.org" Pages 49, 50.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LO_VersionGet($bSimpleVersion = False, $bReturnName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LO_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sAccess = "com.sun.star.configuration.ConfigurationAccess", $sVersionName, $sVersion, $sReturn
	Local $oSettings, $oConfigProvider
	Local $aParamArray[1]

	If Not IsBool($bSimpleVersion) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Local $oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oConfigProvider = $oServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
	If Not IsObj($oConfigProvider) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$aParamArray[0] = __LO_SetPropertyValue("nodepath", "/org.openoffice.Setup/Product")
	If (@error > 0) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	$oSettings = $oConfigProvider.createInstanceWithArguments($sAccess, $aParamArray)

	$sVersionName = $oSettings.getByName("ooName")

	$sVersion = ($bSimpleVersion) ? ($oSettings.getByName("ooSetupVersion")) : ($oSettings.getByName("ooSetupVersionAboutBox"))

	$sReturn = ($bReturnName) ? ($sVersionName & " " & $sVersion) : ($sVersion)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sReturn)
EndFunc   ;==>_LO_VersionGet
