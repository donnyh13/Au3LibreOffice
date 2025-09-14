#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
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
; _LOImpress_ConvertColorFromLong
; _LOImpress_ConvertColorToLong
; _LOImpress_ConvertFromMicrometer
; _LOImpress_ConvertToMicrometer
; _LOImpress_CursorInsertString
; _LOImpress_DrawShapeDelete
; _LOImpress_DrawShapeGetType
; _LOImpress_DrawShapeTextboxCreateTextCursor
; _LOImpress_FontExists
; _LOImpress_FontsGetNames
; _LOImpress_PathConvert
; _LOImpress_VersionGet
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOImpress_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default value is Default. Accepts a Function, or the Keyword Default and Null. If set to a User function, the function may have up to 5 required parameters.
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
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in Autoit HelpFile.) Using the first parameter in the UserFunction.
;                  For Example MyFunc($oMyError)
;                    $oMyError.number The Windows HRESULT value from a COM call
;                    $oMyError.windescription The FormatWinError() text derived from .number
;                    $oMyError.source Name of the Object generating the error (contents from ExcepInfo.source)
;                    $oMyError.description Source Object's description of the error (contents from ExcepInfo.description)
;                    $oMyError.helpfile Source Object's help file for the error (contents from ExcepInfo.helpfile)
;                    $oMyError.helpcontext Source Object's help file context id number (contents from ExcepInfo.helpcontext)
;                    $oMyError.lastdllerror The number returned from GetLastError()
;                    $oMyError.scriptline The script line on which the error was generated
;                    NOTE: Not all properties will necessarily contain data, some will be blank.
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
; Name ..........: _LOImpress_ConvertColorFromLong
; Description ...: Convert Long color code to Hex, RGB, HSB or CMYK.
; Syntax ........: _LOImpress_ConvertColorFromLong([$iHex = Null[, $iRGB = Null[, $iHSB = Null[, $iCMYK = Null]]]])
; Parameters ....: $iHex                - [optional] an integer value. Default is Null. Convert Long Color Integer to Hexadecimal.
;                  $iRGB                - [optional] an integer value. Default is Null. Convert Long Color Integer to R.G.B.
;                  $iHSB                - [optional] an integer value. Default is Null. Convert Long Color Integer to H.S.B.
;                  $iCMYK               - [optional] an integer value. Default is Null. Convert Long Color Integer to C.M.Y.K.
; Return values .: Success: String or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = No parameters set.
;                  @Error 1 @Extended 2 Return 0 = No parameters set to an integer.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Long integer converted To Hexadecimal (as a String). (Without the "0x" prefix)
;                  @Error 0 @Extended 2 Return Array = Array containing Long integer converted To Red, Green, Blue,(RGB). $Array[0] = R, $Array[1] = G, etc.
;                  @Error 0 @Extended 3 Return Array = Array containing Long integer converted To Hue, Saturation, Brightness, (HSB). $Array[0] = H, $Array[1] = S, etc.
;                  @Error 0 @Extended 4 Return Array = Array containing Long integer converted To Cyan, Magenta, Yellow, Black, (CMYK). $Array[0] = C, $Array[1] = M, etc.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve a Hexadecimal color value, call the Long Color code in $iHex, To retrieve a R(ed)G(reen)B(lue) color value, call Null in $iHex, and call the Long color code into $iRGB, etc. for the other color types.
;                  Hex returns as a string variable, all others (RGB, HSB, CMYK) return an array.
;                  The Hexadecimal figure returned doesn't contain the usual "0x", as LibeOffice does not implement it in its numbering system.
; Related .......: _LOImpress_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ConvertColorFromLong($iHex = Null, $iRGB = Null, $iHSB = Null, $iCMYK = Null)
	Local $nRed, $nGreen, $nBlue, $nResult, $nMaxRGB, $nMinRGB, $nHue, $nSaturation, $nBrightness, $nCyan, $nMagenta, $nYellow, $nBlack
	Local $dHex
	Local $aiReturn[0]

	If (@NumParams = 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Select
		Case IsInt($iHex) ; Long to Hex
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
			$nRed = (Mod(($iHSB / 65536), 256)) / 255
			$nGreen = (Mod(($iHSB / 256), 256)) / 255
			$nBlue = (Mod($iHSB, 256)) / 255

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
EndFunc   ;==>_LOImpress_ConvertColorFromLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ConvertColorToLong
; Description ...: Convert Hex, RGB, HSB or CMYK to Long color code.
; Syntax ........: _LOImpress_ConvertColorToLong([$vVal1 = Null[, $vVal2 = Null[, $vVal3 = Null[, $vVal4 = Null]]]])
; Parameters ....: $vVal1               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal2               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal3               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal4               - [optional] a variant value. Default is Null. See remarks.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = No parameters set.
;                  @Error 1 @Extended 2 Return 0 = One parameter called, but not in String format(Hex).
;                  @Error 1 @Extended 3 Return 0 = Hex parameter contains non Hex characters.
;                  @Error 1 @Extended 4 Return 0 = Hex parameter not 6 characters long.
;                  @Error 1 @Extended 5 Return 0 = Hue parameter contains more than just digits.
;                  @Error 1 @Extended 6 Return 0 = Saturation parameter contains more than just digits.
;                  @Error 1 @Extended 7 Return 0 = Brightness parameter contains more than just digits.
;                  @Error 1 @Extended 8 Return 0 = Three parameters called but not all Integers (RGB) and not all Strings (HSB).
;                  @Error 1 @Extended 9 Return 0 = Four parameters called but not all Integers(CMYK).
;                  @Error 1 @Extended 10 Return 0 = Too many or too few parameters called.
;                  --Success--
;                  @Error 0 @Extended 1 Return Integer = Long Int. Color code converted from Hexadecimal.
;                  @Error 0 @Extended 2 Return Integer = Long Int. Color code converted from Red, Green, Blue, (RGB).
;                  @Error 0 @Extended 3 Return Integer = Long Int. Color code converted from (H)ue, (S)aturation, (B)rightness,
;                  @Error 0 @Extended 4 Return Integer = Long Int. Color code converted from (C)yan, (M)agenta, (Y)ellow, Blac(k)
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To Convert a Hex(adecimal) color code, call the Hex code in $vVal1 in String Format.
;                  To convert a R(ed) G(reen) B(lue color, call R value in $vVal1 as an Integer, G in $vVal2 as an Integer, and B in $vVal3 as an Integer.
;                  To convert a H(ue) S(aturation) B(rightness) color, call H in $vVal1 as a String, S in $vVal2 as a String, and B in $vVal3 as a string.
;                  To convert C(yan) M(agenta) Y(ellow) Blac(k) call C in $vVal1 as an Integer, M in $vVal2 as an Integer, Y in $vVal3 as an Integer, and K in $vVal4 as an Integer format.
;                  The Hexadecimal figure entered cannot contain the usual "0x", as LibeOffice does not implement it in its numbering system.
; Related .......: _LOImpress_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ConvertColorToLong($vVal1 = Null, $vVal2 = Null, $vVal3 = Null, $vVal4 = Null) ; RGB = Int, CMYK = Int, HSB = String, Hex = String.
	Local Const $STR_STRIPALL = 8
	Local $iRed, $iGreen, $iBlue, $iLong, $iHue, $iSaturation, $iBrightness
	Local $dHex
	Local $nMaxRGB, $nMinRGB, $nChroma, $nHuePre, $nCyan, $nMagenta, $nYellow, $nBlack

	If (@NumParams = 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch @NumParams
		Case 1 ; Hex
			If Not IsString($vVal1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; not a string

			$vVal1 = StringStripWS($vVal1, $STR_STRIPALL)
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
				$vVal1 = StringStripWS($vVal1, $STR_STRIPALL)
				$vVal2 = StringStripWS($vVal2, $STR_STRIPALL)
				$vVal3 = StringStripWS($vVal3, $STR_STRIPALL) ; Strip WS so I can check string length in HSB conversion.

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
EndFunc   ;==>_LOImpress_ConvertColorToLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ConvertFromMicrometer
; Description ...: Convert from Micrometer to Inch, Centimeter, Millimeter, or Printer's Points.
; Syntax ........: _LOImpress_ConvertFromMicrometer([$nInchOut = Null[, $nCentimeterOut = Null[, $nMillimeterOut = Null[, $nPointsOut = Null]]]])
; Parameters ....: $nInchOut            - [optional] a general number value. Default is Null. The Micrometers to convert to Inches. See remarks.
;                  $nCentimeterOut      - [optional] a general number value. Default is Null. The Micrometers to convert to Centimeters. See remarks.
;                  $nMillimeterOut      - [optional] a general number value. Default is Null. The Micrometers to convert to Millimeters. See remarks.
;                  $nPointsOut          - [optional] a general number value. Default is Null. The Micrometers to convert to Printer's Points. See remarks.
; Return values .: Success: Number
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $nInchOut not a number.
;                  @Error 1 @Extended 2 Return 0 = $nCentimeterOut not a number.
;                  @Error 1 @Extended 3 Return 0 = $nMillimeterOut not a number.
;                  @Error 1 @Extended 4 Return 0 = $nPointsOut not a number.
;                  @Error 1 @Extended 5 Return 0 = No parameters set to other than Null.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting from Micrometers to Inch.
;                  @Error 3 @Extended 2 Return 0 = Error converting from Micrometers to Centimeter.
;                  @Error 3 @Extended 3 Return 0 = Error converting from Micrometers to Millimeter.
;                  @Error 3 @Extended 4 Return 0 = Error converting from Micrometers to Printer's Points.
;                  --Success--
;                  @Error 0 @Extended 1 Return Number = Converted from Micrometers to Inch.
;                  @Error 0 @Extended 2 Return Number = Converted from Micrometers to Centimeter.
;                  @Error 0 @Extended 3 Return Number = Converted from Micrometers to Millimeter.
;                  @Error 0 @Extended 4 Return Number = Converted from Micrometers to Printer's Points.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To skip a parameter, set it to Null.
;                  If you are converting to Inches, place the Micrometers in $nInchOut, if converting to Millimeters, $nInchOut and $nCentimeter are set to Null, and $nMillimetersOut is set.
;                  A Micrometer is 1000th of a centimeter, and is used in almost all Libre Office functions that contain a measurement parameter.
; Related .......: _LOImpress_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ConvertFromMicrometer($nInchOut = Null, $nCentimeterOut = Null, $nMillimeterOut = Null, $nPointsOut = Null)
	Local $nReturnValue

	If ($nInchOut <> Null) Then
		If Not IsNumber($nInchOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$nReturnValue = __LOImpress_UnitConvert($nInchOut, $__LOCONST_CONVERT_UM_INCH)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterOut <> Null) Then
		If Not IsNumber($nCentimeterOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$nReturnValue = __LOImpress_UnitConvert($nCentimeterOut, $__LOCONST_CONVERT_UM_CM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterOut <> Null) Then
		If Not IsNumber($nMillimeterOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$nReturnValue = __LOImpress_UnitConvert($nMillimeterOut, $__LOCONST_CONVERT_UM_MM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Return SetError($__LO_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsOut <> Null) Then
		If Not IsNumber($nPointsOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$nReturnValue = __LOImpress_UnitConvert($nPointsOut, $__LOCONST_CONVERT_UM_PT)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.
EndFunc   ;==>_LOImpress_ConvertFromMicrometer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ConvertToMicrometer
; Description ...: Convert from Inch, Centimeter, Millimeter, or Printer's Points to Micrometer.
; Syntax ........: _LOImpress_ConvertToMicrometer([$nInchIn = Null[, $nCentimeterIn = Null[, $nMillimeterIn = Null[, $nPointsIn = Null]]]])
; Parameters ....: $nInchIn             - [optional] a general number value. Default is Null. The Inches to convert to Micrometers. See remarks.
;                  $nCentimeterIn       - [optional] a general number value. Default is Null. The Centimeters to convert to Micrometers. See remarks.
;                  $nMillimeterIn       - [optional] a general number value. Default is Null. The Millimeters to convert to Micrometers. See remarks.
;                  $nPointsIn           - [optional] a general number value. Default is Null. The Printer's Points to convert to Micrometers. See remarks.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $nInchIn not a number.
;                  @Error 1 @Extended 2 Return 0 = $nCentimeterIn not a number.
;                  @Error 1 @Extended 3 Return 0 = $nMillimeterIn not a number.
;                  @Error 1 @Extended 4 Return 0 = $nPointsIn not a number.
;                  @Error 1 @Extended 5 Return 0 = No parameters set to other than Null.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error converting from Inches to Micrometers.
;                  @Error 3 @Extended 2 Return 0 = Error converting from Centimeters to Micrometers.
;                  @Error 3 @Extended 3 Return 0 = Error converting from Millimeters to Micrometers.
;                  @Error 3 @Extended 4 Return 0 = Error converting from Printer's Points to Micrometers.
;                  --Success--
;                  @Error 0 @Extended 1 Return Integer = Converted Inches to Micrometers.
;                  @Error 0 @Extended 2 Return Integer = Converted Centimeters to Micrometers.
;                  @Error 0 @Extended 3 Return Integer = Converted Millimeters to Micrometers.
;                  @Error 0 @Extended 4 Return Integer = Converted Printer's Points to Micrometers.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To skip a parameter, set it to Null.
;                  If you are converting from Inches, call inches in $nInchIn, if converting from Centimeters, $nInchIn is called with Null, and $nCentimeters is set.
;                  A Micrometer is 1000th of a centimeter, and is used in almost all Libre Office functions that contain a measurement parameter.
; Related .......: _LOImpress_ConvertFromMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ConvertToMicrometer($nInchIn = Null, $nCentimeterIn = Null, $nMillimeterIn = Null, $nPointsIn = Null)
	Local $nReturnValue

	If ($nInchIn <> Null) Then
		If Not IsNumber($nInchIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		$nReturnValue = __LOImpress_UnitConvert($nInchIn, $__LOCONST_CONVERT_INCH_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterIn <> Null) Then
		If Not IsNumber($nCentimeterIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$nReturnValue = __LOImpress_UnitConvert($nCentimeterIn, $__LOCONST_CONVERT_CM_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		Return SetError($__LO_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterIn <> Null) Then
		If Not IsNumber($nMillimeterIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$nReturnValue = __LOImpress_UnitConvert($nMillimeterIn, $__LOCONST_CONVERT_MM_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		Return SetError($__LO_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsIn <> Null) Then
		If Not IsNumber($nPointsIn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$nReturnValue = __LOImpress_UnitConvert($nPointsIn, $__LOCONST_CONVERT_PT_UM)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

		Return SetError($__LO_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.
EndFunc   ;==>_LOImpress_ConvertToMicrometer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorInsertString
; Description ...: Insert a string at a cursor position.
; Syntax ........: _LOImpress_CursorInsertString(ByRef $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oCursor             - [in/out] an object. A Text Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $sString             - a string value. A String to insert.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, and the cursor object has text selected, the selection is overwritten, else if False, the string is inserted to the left of the selection. If there are multiple selections, the string is inserted to the left of the last selection, and none are overwritten.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sString not a string..
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. String was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Warning! For some reason this function doesn't seem to set the modified status to True. Changes could be inadvertently lost due to this, if the user closes without saving.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorInsertString(ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCursor.Text.insertString($oCursor, $sString, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_CursorInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeDelete
; Description ...: Delete a Shape.
; Syntax ........: _LOImpress_DrawShapeDelete(ByRef $oShape)
; Parameters ....: $oShape                - [in/out] an object. A Shape object returned by a previous _LOImpress_ShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape's containing Slide.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of shapes.
;                  @Error 3 @Extended 3 Return 0 = Same number of shapes still present. Failed to delete the Shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Shape was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeDelete(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iShapes
	Local $oDrawPage

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDrawPage = $oShape.Parent()
	If Not IsObj($oDrawPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iShapes = $oDrawPage.getCount()
	If Not IsInt($iShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDrawPage.remove($oShape)

	Return ($oDrawPage.getCount() = $iShapes) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeGetType
; Description ...: Return the Drawing Shape's Type corresponding to the constants $LOI_DRAWSHAPE_TYPE_*
; Syntax ........: _LOImpress_DrawShapeGetType(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_ShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve CustomShapeGeometry Array.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve CustomShapeGeometry "Type" value.
;                  @Error 3 @Extended 3 Return 0 = Failed to determine CustomShape's type.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify what type of "com.sun.star.drawing.EllipseShape" called shape is.
;                  @Error 3 @Extended 5 Return 0 = Called Shape is a unknown shape type.
;                  --Success--
;                  @Error 0 @Extended 1 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_CONNECTOR_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_CONNECTOR Constant Value. See Remarks #4.
;                  @Error 0 @Extended 2 Return Integer = Success. Shape is a Custom Shape Type. Returning appropriate Constant for shape type if successfully identified, else -1 if identification failed. See Remarks #1. See Constants, $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 3 Return Integer = Success. Shape is a*_BASIC_CIRCLE_SEGMENT or *_BASIC_ARC Type Shape. Returning appropriate Constant, See Constants, $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 4 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_CURVE Shape.
;                  @Error 0 @Extended 5 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED Shape.
;                  @Error 0 @Extended 6 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE Shape.
;                  @Error 0 @Extended 7 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED Shape.
;                  @Error 0 @Extended 8 Return Integer = Success. Shape is a *_LINE_LINE, *_LINE_LINE_45 or $LOI_DRAWSHAPE_TYPE_LINE_ARROW_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_LINE Constant Value. See Remarks #4.
;                  @Error 0 @Extended 9 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION Shape.
;                  @Error 0 @Extended 10 Return Integer = Success. Shape is a *_LINE_POLYGON, or *_LINE_POLYGON_45 Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_POLYGON Constant Value. See Remarks #2.
;                  @Error 0 @Extended 11 Return Integer = Success. Shape is a *_LINE_POLYGON_FILLED, or *_LINE_POLYGON_45_FILLED Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED Constant Value. See Remarks #2.
;                  @Error 0 @Extended 11 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_3D_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_3D_CONE Constant Value. See Remarks #5.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: #1 Some shapes are not implemented, or not fully implemented into LibreOffice for automation, consequently they do not have appropriate type names as of yet. Many have simply ambiguous names, such as "non-primitive".
;                  Because of this the following Custom shape types cannot be identified, and this function will return -1:
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, known as "mso-spt100".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, known as "non-primitive", should be "corner-right-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, known as "non-primitive", should be "right-left-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, known as "non-primitive", should be "s-sharped-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, known as "non-primitive", should be "split-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, known as "mso-spt100", should be "striped-right-arrow".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, known as "mso-spt89", should be "up-right-arrow-callout".
;                   $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, known as "mso-spt100", should be "up-right-down-arrow".
;                   $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, known as "mso-spt100", should be "circle-pie".
;                   $LOI_DRAWSHAPE_TYPE_STARS_6_POINT, known as "non-primitive", should be "star6".
;                   $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE, known as "non-primitive", should be "concave-star6".
;                   $LOI_DRAWSHAPE_TYPE_STARS_12_POINT, known as "non-primitive", should be "star12".
;                   $LOI_DRAWSHAPE_TYPE_STARS_SIGNET, known as "non-primitive", should be "signet".
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, known as "non-primitive", should be "cloud"?
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, known as "non-primitive", should be "flower"?
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING, known as "non-primitive", should be "lightning".
;                  #2 The following Shapes implement the same type names, and are consequently indistinguishable:
;                   $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE, $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE is returned for either one.)
;                   $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE is returned for either one.)
;                   $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED is returned for either one.)
;                   $LOI_DRAWSHAPE_TYPE_LINE_POLYGON, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45 (The Value of $LOI_DRAWSHAPE_TYPE_LINE_POLYGON is returned for either of these.)
;                   $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED (The Value of $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED is returned for either of these.)
;                  #3 The following Shapes have strange names that may change in the future, but currently are able to be identified:
;                   $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, known as, "mso-spt21", should be "doorplate"
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, known as, "col-502ad400", should be ??
;                   $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, known as, "col-60da8460", should be ??
;                  #4 The following Shapes are customizable one to another, and are consequently indistinguishable:
;                   $LOI_DRAWSHAPE_TYPE_CONNECTOR_* (The Value of $LOI_DRAWSHAPE_TYPE_CONNECTOR is returned for any of these.)
;                   $LOI_DRAWSHAPE_TYPE_FONTWORK_* (The Value of $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL is returned for any of these.)
;                   $LOI_DRAWSHAPE_TYPE_LINE_ARROW_* or $LOI_DRAWSHAPE_TYPE_LINE_LINE_45 (The Value of $LOI_DRAWSHAPE_TYPE_LINE_LINE is returned for any of these.)
;                  #5 The following Shapes are have nothing unique that I have found yet to identify each, and are consequently indistinguishable:
;                   $LOI_DRAWSHAPE_TYPE_3D_* (The Value of $LOI_DRAWSHAPE_TYPE_3D_CONE is returned for any of these.)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeGetType(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atCusShapeGeo[0]
	Local Const $iCircleKind_CUT = 2 ; a circle with a cut connected by a line.
	Local Const $iCircleKind_ARC = 3 ; a circle with an open cut.
	Local $sType
	Local $iReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $oShape.ShapeType()

		Case "com.sun.star.drawing.ConnectorShape" ; No way to differentiate between these??

			Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR)

		Case "com.sun.star.drawing.CustomShape"
			$atCusShapeGeo = $oShape.CustomShapeGeometry()
			If Not IsArray($atCusShapeGeo) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			For $i = 0 To UBound($atCusShapeGeo) - 1
				If ($atCusShapeGeo[$i].Name() = "Type") Then
					$sType = $atCusShapeGeo[$i].Value()
					If Not IsString($sType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

					ExitLoop
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$iReturn = __LOImpress_DrawShape_GetCustomType($sType)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			Return SetError($__LO_STATUS_SUCCESS, 2, $iReturn)

		Case "com.sun.star.drawing.EllipseShape"
			If ($oShape.CircleKind() = $iCircleKind_CUT) Then ; Circle Segment = CircleKind_CUT(2), Arc = CircleKind_ARC(3)

				Return SetError($__LO_STATUS_SUCCESS, 3, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT)

			ElseIf ($oShape.CircleKind() = $iCircleKind_ARC) Then

				Return SetError($__LO_STATUS_SUCCESS, 3, $LOI_DRAWSHAPE_TYPE_BASIC_ARC)

			Else

				Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
			EndIf

		Case "com.sun.star.drawing.OpenBezierShape"

			Return SetError($__LO_STATUS_SUCCESS, 4, $LOI_DRAWSHAPE_TYPE_LINE_CURVE)

		Case "com.sun.star.drawing.ClosedBezierShape"

			Return SetError($__LO_STATUS_SUCCESS, 5, $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED)

		Case "com.sun.star.drawing.OpenFreeHandShape"

			Return SetError($__LO_STATUS_SUCCESS, 6, $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE)

		Case "com.sun.star.drawing.ClosedFreeHandShape"

			Return SetError($__LO_STATUS_SUCCESS, 7, $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED)

		Case "com.sun.star.drawing.LineShape" ; No way to differentiate between these?? (Lines + Arrows)

			Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_LINE)

		Case "com.sun.star.drawing.MeasureShape"

			Return SetError($__LO_STATUS_SUCCESS, 9, $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION)

		Case "com.sun.star.drawing.PolyLineShape"
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON ; No way to differentiate between these??
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45

			Return SetError($__LO_STATUS_SUCCESS, 10, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON)

		Case "com.sun.star.drawing.PolyPolygonShape"

			Return SetError($__LO_STATUS_SUCCESS, 11, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED)
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED ; No way to differentiate between these??
;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED

		Case "com.sun.star.drawing.Shape3DSceneObject" ; No way to differentiate between these??

			Return SetError($__LO_STATUS_SUCCESS, 12, $LOI_DRAWSHAPE_TYPE_3D_CONE)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Unknown shape type.
	EndSwitch
EndFunc   ;==>_LOImpress_DrawShapeGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeTextboxCreateTextCursor
; Description ...: Create a Text Cursor in a Textbox for inserting text etc.
; Syntax ........: _LOImpress_DrawShapeTextboxCreateTextCursor(ByRef $oTextbox)
; Parameters ....: $oTextbox              - [in/out] an object. A Frame object returned by a previous _LOImpress_FrameCreate, _LOImpress_FrameGetObjByName, or _LOImpress_FrameGetObjByCursor function.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextbox not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. A Text Cursor Object located in the Frame.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeTextboxCreateTextCursor(ByRef $oTextbox)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor
	Local $iShapeType

	If Not IsObj($oTextbox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iShapeType = __LOImpress_DrawShapeGetType($oTextbox)
	If ($iShapeType <> $LOI_SHAPE_TYPE_TEXTBOX) And ($iShapeType <> $LOI_SHAPE_TYPE_TEXTBOX_TITLE) And ($iShapeType <> $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextCursor = $oTextbox.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOImpress_DrawShapeTextboxCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FontExists
; Description ...: Tests whether a specific font exists by name.
; Syntax ........: _LOImpress_FontExists($sFontName[, $oDoc = Null])
; Parameters ....: $sFontName           - a string value. The Font name to search for.
;                  $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
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
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the Font is available, else False.
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

		$atProperties[0] = __LOImpress_SetPropertyValue("Hidden", True)
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
; Parameters ....: $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
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
;                  @Error 0 @Extended ? Return Array = Success. Returns a 4 Column Array, @extended is set to the number of results. See remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, an Impress Document is created invisibly to perform the check.
;                  Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold, Italic, etc. Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for easier processing if required.
;                  From personal tests, Slant only returns 0 or 2.
;                  The returned array will be as follows:
;                  The first column (Array[1][0]) contains the Font Name.
;                  The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOI_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOI_POSTURE_* as defined in LibreOfficeImpress_Constants.au3.
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

		$atProperties[0] = __LOImpress_SetPropertyValue("Hidden", True)
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

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_PathConvert
; Description ...: Converts the input path to or from a LibreOffice URL notation path.
; Syntax ........: _LOImpress_PathConvert($sFilePath[, $iReturnMode = $LOI_PATHCONV_AUTO_RETURN])
; Parameters ....: $sFilePath           - a string value. Full path to convert in String format.
;                  $iReturnMode         - [optional] an integer value (0-2). Default is $__g_iAutoReturn. The type of path format to return. See Constants, $LOI_PATHCONV_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFilePath is not a string
;                  @Error 1 @Extended 2 Return 0 = $iReturnMode not a Integer, less than 0, or greater than 2, see constants, $LOI_PATHCONV_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Returning converted File Path from Libre Office URL.
;                  @Error 0 @Extended 2 Return String = Returning converted path from File Path to Libre Office URL.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice URL notation is based on the Internet Standard RFC 1738, which means only [0-9],[a-zA-Z] are allowed in paths, most other characters need to be converted into ISO 8859-1 (ISO Latin) such as is found in internet URL's (spaces become %20). See: StarOfficeTM 6.0 Office SuiteA SunTM ONE Software Offering, Basic Programmer's Guide; Page 74
;                  The user generally should not even need this function, as I have endeavored to convert any URLs to the appropriate computer path format and any input computer paths to a Libre Office URL.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_PathConvert($sFilePath, $iReturnMode = $LOI_PATHCONV_AUTO_RETURN)
	Local Const $STR_STRIPLEADING = 1
	Local $asURLReplace[9][2] = [["%", "%25"], [" ", "%20"], ["\", "/"], [";", "%3B"], ["#", "%23"], ["^", "%5E"], ["{", "%7B"], ["}", "%7D"], ["`", "%60"]]
	Local $iPathSearch, $iFileSearch, $iPartialPCPath, $iPartialFilePath

	If Not IsString($sFilePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOImpress_IntIsBetween($iReturnMode, $LOI_PATHCONV_AUTO_RETURN, $LOI_PATHCONV_PCPATH_RETURN) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sFilePath = StringStripWS($sFilePath, $STR_STRIPLEADING)

	$iPathSearch = StringRegExp($sFilePath, "[A-Z]\:\\") ; Search For a Computer Path, as in C:\ etc.
	$iPartialPCPath = StringInStr($sFilePath, "\") ; Search for partial computer Path containing a backslash.
	$iFileSearch = StringInStr($sFilePath, "file:///", 0, 1, 1, 9) ; Search for a full Libre path, which begins with File:///
	$iPartialFilePath = StringInStr($sFilePath, "/") ; Search For a Partial Libre path containing forward slash

	If ($iReturnMode = $LOI_PATHCONV_AUTO_RETURN) Then
		If ($iPathSearch > 0) Or ($iPartialPCPath > 0) Then ;  if file path contains partial or full PC path, set to convert to Libre URL.
			$iReturnMode = $LOI_PATHCONV_OFFICE_RETURN

		ElseIf ($iFileSearch > 0) Or ($iPartialFilePath > 0) Then ;  if file path contains partial or full Libre URL, set to convert to PC Path.
			$iReturnMode = $LOI_PATHCONV_PCPATH_RETURN

		Else ; If file path contains neither above. convert to Libre URL
			$iReturnMode = $LOI_PATHCONV_OFFICE_RETURN
		EndIf
	EndIf

	Switch $iReturnMode
		Case $LOI_PATHCONV_OFFICE_RETURN
			If $iFileSearch > 0 Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)
			If ($iPathSearch > 0) Then $sFilePath = "file:///" & $sFilePath

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][0], $asURLReplace[$i][1])
				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			Return SetError($__LO_STATUS_SUCCESS, 2, $sFilePath)

		Case $LOI_PATHCONV_PCPATH_RETURN
			If ($iPathSearch > 0) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)
			If ($iFileSearch > 0) Then $sFilePath = StringReplace($sFilePath, "file:///", Null)

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][1], $asURLReplace[$i][0])
				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			Return SetError($__LO_STATUS_SUCCESS, 1, $sFilePath)
	EndSwitch
EndFunc   ;==>_LOImpress_PathConvert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_VersionGet
; Description ...: Retrieve the current Office version.
; Syntax ........: _LOImpress_VersionGet([$bSimpleVersion = False[, $bReturnName = False]])
; Parameters ....: $bSimpleVersion      - [optional] a boolean value. Default is False. If True, returns a two digit version number, such as "7.3", else returns the complex version number, such as "7.3.2.4".
;                  $bReturnName         - [optional] a boolean value. Default is True. If True returns the Program Name, such as "LibreOffice", appended by the version, i.e. "LibreOffice 7.3".
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bSimpleVersion not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $bReturnName not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.configuration.ConfigurationProvider" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error setting property value.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Returns the Office version in String format.
; Author ........: Laurent Godard as found in Andrew Pitonyak's book; Zizi64 as found on OpenOffice forum.
; Modified ......: donnyh13, modified for AutoIt compatibility and error checking.
; Remarks .......: From Macro code by Zizi64 found at: https://forum.openoffice.org/en/forum/viewtopic.php?t=91542&sid=7f452d65e58ac1cd3cc6063350b5ada0
;                  And Andrew Pitonyak in "Useful Macro Information For OpenOffice.org" Pages 49, 50.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_VersionGet($bSimpleVersion = False, $bReturnName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
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

	$aParamArray[0] = __LOImpress_SetPropertyValue("nodepath", "/org.openoffice.Setup/Product")
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oSettings = $oConfigProvider.createInstanceWithArguments($sAccess, $aParamArray)

	$sVersionName = $oSettings.getByName("ooName")

	$sVersion = ($bSimpleVersion) ? ($oSettings.getByName("ooSetupVersion")) : ($oSettings.getByName("ooSetupVersionAboutBox"))

	$sReturn = ($bReturnName) ? ($sVersionName & " " & $sVersion) : ($sVersion)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sReturn)
EndFunc   ;==>_LOImpress_VersionGet
