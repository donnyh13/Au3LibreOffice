#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Internal.au3"
#include "LibreOfficeWriter_Font.au3"
#include "LibreOfficeWriter_Page.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13, mLipok
; Sources .......: jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
;					mLipok -- OOoCalc.au3, used (__OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler,
;						-- WriterDemo.au3, used _CreateStruct;
;					Andrew Pitonyak & Laurent Godard (VersionGet);
;					Leagnus & GMK -- OOoCalc.au3, used (SetPropertyValue)
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;					I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained;
;						OOME Third Edition".
;					Of course, this UDF is written using the English version of LibreOffice, and may only work for the English
;						version of LibreOffice installations. Many functions in this UDF may or may not work with OpenOffice
;						Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_ComError_UserFunction
; _LOWriter_ConvertColorFromLong
; _LOWriter_ConvertColorToLong
; _LOWriter_ConvertFromMicrometer
; _LOWriter_ConvertToMicrometer
; _LOWriter_DateFormatKeyCreate
; _LOWriter_DateFormatKeyDelete
; _LOWriter_DateFormatKeyExists
; _LOWriter_DateFormatKeyGetString
; _LOWriter_DateFormatKeyList
; _LOWriter_DateStructCreate
; _LOWriter_DateStructModify
; _LOWriter_FindFormatModifyAlignment
; _LOWriter_FindFormatModifyEffects
; _LOWriter_FindFormatModifyFont
; _LOWriter_FindFormatModifyHyphenation
; _LOWriter_FindFormatModifyIndent
; _LOWriter_FindFormatModifyOverline
; _LOWriter_FindFormatModifyPageBreak
; _LOWriter_FindFormatModifyPosition
; _LOWriter_FindFormatModifyRotateScaleSpace
; _LOWriter_FindFormatModifySpacing
; _LOWriter_FindFormatModifyStrikeout
; _LOWriter_FindFormatModifyTxtFlowOpt
; _LOWriter_FindFormatModifyUnderline
; _LOWriter_FormatKeyCreate
; _LOWriter_FormatKeyDelete
; _LOWriter_FormatKeyExists
; _LOWriter_FormatKeyGetString
; _LOWriter_FormatKeyList
; _LOWriter_PathConvert
; _LOWriter_VersionGet
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOWriter_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default value is Default. Accepts a Function, or the Keyword Default and Null. If set to a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = In correct entry. Not a Function, or Default keyword or Null Keyword.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Successfully set the UserFunction.
;				   @Error 0 @Extended 0 Return 2 = Successfully cleared the set UserFunction.
;				   @Error 0 @Extended 0 Return Function = Returns the set UserFunction.
; Author ........: mLipok
; Modified ......: donnyh13 - Added a clear UserFunction without error option. Also added parameters option.
; Remarks .......: The first parameter passed to the User function will always be the COM Error object. See below.
;						Every COM Error will be passed to that function. The user can then read the following properties. (As
;							Found in the COM Reference section in Autoit HelpFile.) Using the first parameter in the
;							UserFunction. For Example MyFunc($oMyError)
;							$oMyError.number The Windows HRESULT value from a COM call
;							$oMyError.windescription The FormatWinError() text derived from .number
;							$oMyError.source Name of the Object generating the error (contents from ExcepInfo.source)
;							$oMyError.description Source Object's description of the error (contents from ExcepInfo.description)
;							$oMyError.helpfile Source Object's helpfile for the error (contents from ExcepInfo.helpfile)
;							$oMyError.helpcontext Source Object's helpfile context id number (contents from ExcepInfo.helpcontext)
;							$oMyError.lastdllerror The number returned from GetLastError()
;							$oMyError.scriptline The script line on which the error was generated
;				    		NOTE: Not all properties will necessarily contain data, some will be blank.
;						If MsgBox or ConsoleWrite functions are passed to this function, the error details will be displayed using that function automatically.
;						If called with Default keyword, the current UserFunction, if set, will be returned.
;				    	If called with Null keyword, the currently set UserFunction is cleared and only the internal ComErrorHandler will be called for COM Errors.
;						The stored UserFunction (besides MsgBox and ConsoleWrite) will be called as follows: UserFunc($oComError,$vParam1,$vParam2,$vParam3,$vParam4,$vParam5)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
	#forceref $vParam1, $vParam2, $vParam3, $vParam4, $vParam5
	; If user does not set a function, UDF must use internal function to avoid AutoItError.
	Local Static $vUserFunction_Static = Default
	Local $avUserFuncWParams[@NumParams]

	If $vUserFunction = Default Then
		; just return stored static User Function variable
		Return $vUserFunction_Static
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
		Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
	ElseIf $vUserFunction = Null Then
		; Clear User Function.
		$vUserFunction_Static = Default
		Return SetError($__LOW_STATUS_SUCCESS, 0, 2)
	Else
		; return error as an incorrect parameter was passed to this function
		Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOWriter_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ConvertColorFromLong
; Description ...: Convert Long color code to Hex, RGB, HSB or CMYK.
; Syntax ........: _LOWriter_ConvertColorFromLong([$iHex = Null[, $iRGB = Null[, $iHSB = Null[, $iCMYK = Null]]]])
; Parameters ....: $iHex                - [optional] an integer value. Default is Null.
;                  $iRGB                - [optional] an integer value. Default is Null.
;                  $iHSB                - [optional] an integer value. Default is Null.
;                  $iCMYK               - [optional] an integer value. Default is Null.
; Return values .: Success: Hex or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = No parameters set.
;				   @Error 1 @Extended 2 Return 0 = No parameters set to an integer.
;				   --Success--
;				   @Error 0 @Extended 1 Return Hex Color Integer. Long integer converted To Hexadecimal. (Without the "0x" prefix)
;				   @Error 0 @Extended 2 Return Array. Array containing Long integer converted To Red, Green, Blue,(RGB). $Array[3] = [R,G,B] = $Array[0] = R, etc.
;				   @Error 0 @Extended 3 Return Array. Array containing Long integer converted To Hue, Saturation, Brightness, (HSB). $Array[3] = [H,S,B] $Array[0] = H, etc.
;				   @Error 0 @Extended 4 Return Array. Array containing Long integer converted To Cyan, Yellow, Magenta, Black, (CMYK). $Array[4] = [C,M,Y,K] $Array[0] = C, etc.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve a Hex(adecimal) color code, place the Long Color code into $iHex, To retrieve a R(ed)G(reen)B(lue)
;					color code, place Null in $iHex, and the Long color code into $iRGB, etc. for the other color types.
;					Hex returns as a string variable, all others (RGB, HSB, CMYK) return an array. Array[0] = R, Array [1] = G etc.
;					Note: The Hexadecimal figure returned doesn't contain the usual "0x", as LibeOffice does not implement it in its numbering system.
; Related .......: _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ConvertColorFromLong($iHex = Null, $iRGB = Null, $iHSB = Null, $iCMYK = Null)
	Local $nRed, $nGreen, $nBlue, $nResult, $nMaxRGB, $nMinRGB, $nHue, $nSaturation, $nBrightness, $nCyan, $nMagenta, $nYellow, $nBlack
	Local $dHex
	Local $aiReturn[0]

	If (@NumParams = 0) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	Select
		Case IsInt($iHex) ; Long TO Hex
			$nRed = Int(Mod(($iHex / 65536), 256))
			$nGreen = Int(Mod(($iHex / 256), 256))
			$nBlue = Int(Mod($iHex, 256))

			$dHex = Hex($nRed, 2) & Hex($nGreen, 2) & Hex($nBlue, 2)
			Return SetError($__LOW_STATUS_SUCCESS, 1, $dHex)

		Case IsInt($iRGB) ; Long to RGB
			$nRed = Int(Mod(($iRGB / 65536), 256))
			$nGreen = Int(Mod(($iRGB / 256), 256))
			$nBlue = Int(Mod($iRGB, 256))
			ReDim $aiReturn[3]
			$aiReturn[0] = $nRed
			$aiReturn[1] = $nGreen
			$aiReturn[2] = $nBlue
			Return SetError($__LOW_STATUS_SUCCESS, 2, $aiReturn)
		Case IsInt($iHSB) ; Long TO HSB

			$nRed = (Mod(($iHSB / 65536), 256)) / 255
			$nGreen = (Mod(($iHSB / 256), 256)) / 255
			$nBlue = (Mod($iHSB, 256)) / 255

			; get Max RGB Value
			$nResult = ($nRed > $nGreen) ? $nRed : $nGreen
			$nMaxRGB = ($nResult > $nBlue) ? $nResult : $nBlue
			; get Min RGB Value
			$nResult = ($nRed < $nGreen) ? $nRed : $nGreen
			$nMinRGB = ($nResult < $nBlue) ? $nResult : $nBlue

			; Determine Brightness
			$nBrightness = $nMaxRGB
			; Determine Hue
			$nHue = 0
			Select
				Case $nRed = $nGreen = $nBlue ; Red, Green, and BLue are equal.
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
			$nSaturation = ($nMaxRGB = 0) ? 0 : (($nMaxRGB - $nMinRGB) / $nMaxRGB)

			$nHue = ($nHue > 0) ? Round($nHue) : 0
			$nSaturation = Round(($nSaturation * 100))
			$nBrightness = Round(($nBrightness * 100))

			ReDim $aiReturn[3]
			$aiReturn[0] = $nHue
			$aiReturn[1] = $nSaturation
			$aiReturn[2] = $nBrightness

			Return SetError($__LOW_STATUS_SUCCESS, 3, $aiReturn)
		Case IsInt($iCMYK) ; Long to CMYK

			$nRed = (Mod(($iCMYK / 65536), 256))
			$nGreen = (Mod(($iCMYK / 256), 256))
			$nBlue = (Mod($iCMYK, 256))

			$nRed = Round(($nRed / 255), 3)
			$nGreen = Round(($nGreen / 255), 3)
			$nBlue = Round(($nBlue / 255), 3)

			; get Max RGB Value
			$nResult = ($nRed > $nGreen) ? $nRed : $nGreen
			$nMaxRGB = ($nResult > $nBlue) ? $nResult : $nBlue

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
			Return SetError($__LOW_STATUS_SUCCESS, 4, $aiReturn)
		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; no parameters set to an integer
	EndSelect

EndFunc   ;==>_LOWriter_ConvertColorFromLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ConvertColorToLong
; Description ...: Convert Hex, RGB, HSB or CMYK to Long color code.
; Syntax ........: _LOWriter_ConvertColorToLong([$vVal1 = Null[, $vVal2 = Null[, $vVal3 = Null[, $vVal4 = Null]]]])
; Parameters ....: $vVal1               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal2               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal3               - [optional] a variant value. Default is Null. See remarks.
;                  $vVal4               - [optional] a variant value. Default is Null. See remarks.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = No parameters set.
;				   @Error 1 @Extended 2 Return 0 = One parameter called, but not in String format(Hex).
;				   @Error 1 @Extended 3 Return 0 = Hex parameter contains non Hex characters.
;				   @Error 1 @Extended 4 Return 0 = Hex parameter not 6 characters long.
;				   @Error 1 @Extended 5 Return 0 = Hue parameter contains more than just digits.
;				   @Error 1 @Extended 6 Return 0 = Saturation parameter contains more than just digits.
;				   @Error 1 @Extended 7 Return 0 = Brightness parameter contains more than just digits.
;				   @Error 1 @Extended 8 Return 0 = Three parameters called but not all Integers (RGB) and not all Strings (HSB).
;				   @Error 1 @Extended 9 Return 0 = Four parameters called but not all Integers(CMYK).
;				   @Error 1 @Extended 10 Return 0 = Too many or too few parameters called.
;				   --Success--
;				   @Error 0 @Extended 1 Return Integer. Long Int. Color code converted from Hexadecimal.
;				   @Error 0 @Extended 2 Return Integer. Long Int. Color code converted from Red, Green, Blue, (RGB).
;				   @Error 0 @Extended 3 Return Integer. Long Int. Color code converted from (H)ue, (S)aturation, (B)rightness,
;				   @Error 0 @Extended 4 Return Integer. Long Int. Color code converted from (C)yan, (Y)ellow, (M)agenta, (B)lack
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To Convert a Hex(adecimal) color code, insert the Hex code into $vVal1 in String Format.
;					To convert a R(ed)G(reen)B(lue color code, insert R in $vVal1 in Integer format, G in $vVal2 in Integer
;					format, and B in $vVal3 in Integer format.
;					To convert H(ue)S(aturation)B(rightness) color code, insert H in $vVal1 in String format, B in $vVal2 in String format, and B in $vVal3 in string format.
;					To convert C(yan)Y(ellow)M(agenta)(Blac)K enter C in $vVal1 in Integer format, Y in $vVal2 in Integer Format, M in $vVal3 in Integer format, and K in $vVal4 in Integer format.
;					Note: The Hexadecimal figure entered cannot contain the usual "0x", as LibeOffice does not implement it in its numbering system.
; Related .......: _LOWriter_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ConvertColorToLong($vVal1 = Null, $vVal2 = Null, $vVal3 = Null, $vVal4 = Null) ; RGB = Int, CMYK = Int, HSB = String, Hex = String.
	Local Const $STR_STRIPALL = 8
	Local $iRed, $iGreen, $iBlue, $iLong, $iHue, $iSaturation, $iBrightness
	Local $dHex
	Local $nMaxRGB, $nMinRGB, $nChroma, $nHuePre, $nCyan, $nMagenta, $nYellow, $nBlack

	If (@NumParams = 0) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	Switch @NumParams
		Case 1 ;Hex
			If Not IsString($vVal1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; not a string
			$vVal1 = StringStripWS($vVal1, $STR_STRIPALL)
			$dHex = $vVal1

			; From Hex to RGB
			If (StringLen($dHex) = 6) Then
				If StringRegExp($dHex, "[^0-9a-fA-F]") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; $dHex contains non Hex characters.

				$iRed = BitAND(BitShift("0x" & $dHex, 16), 0xFF)
				$iGreen = BitAND(BitShift("0x" & $dHex, 8), 0xFF)
				$iBlue = BitAND("0x" & $dHex, 0xFF)

				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue
				Return SetError($__LOW_STATUS_SUCCESS, 1, $iLong) ; Long from Hex

			Else
				Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0) ; Wrong length of string.
			EndIf

		Case 3 ;RGB and HSB; HSB is all strings, RGB all Integers.
			If (IsInt($vVal1) And IsInt($vVal2) And IsInt($vVal3)) Then ; RGB
				$iRed = $vVal1
				$iGreen = $vVal2
				$iBlue = $vVal3

				; RGB to Long
				$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue
				Return SetError($__LOW_STATUS_SUCCESS, 2, $iLong) ; Long from RGB

			ElseIf IsString($vVal1) And IsString($vVal2) And IsString($vVal3) Then ; Hue Saturation and Brightness (HSB)

				; HSB to RGB
				$vVal1 = StringStripWS($vVal1, $STR_STRIPALL)
				$vVal2 = StringStripWS($vVal2, $STR_STRIPALL)
				$vVal3 = StringStripWS($vVal3, $STR_STRIPALL) ; Strip WS so I can check string length in HSB conversion.

				$iHue = Number($vVal1)
				If (StringLen($vVal1)) <> (StringLen($iHue)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ; String contained more than just digits
				$iSaturation = Number($vVal2)
				If (StringLen($vVal2)) <> (StringLen($iSaturation)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0) ; String contained more than just digits
				$iBrightness = Number($vVal3)
				If (StringLen($vVal3)) <> (StringLen($iBrightness)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0) ; String contained more than just digits

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
				Return SetError($__LOW_STATUS_SUCCESS, 3, $iLong) ; Return Long from HSB
			Else
				Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0) ; Wrong parameters
			EndIf
		Case 4 ;CMYK
			If Not (IsInt($vVal1) And IsInt($vVal2) And IsInt($vVal3) And IsInt($vVal4)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0) ; CMYK not integers.

			; CMYK to RGB
			$nCyan = ($vVal1 / 100)
			$nMagenta = ($vVal2 / 100)
			$nYellow = ($vVal3 / 100)
			$nBlack = ($vVal4 / 100)

			$iRed = Round((255 * (1 - $nBlack) * (1 - $nCyan)))
			$iGreen = Round((255 * (1 - $nBlack) * (1 - $nMagenta)))
			$iBlue = Round((255 * (1 - $nBlack) * (1 - $nYellow)))

			$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue
			Return SetError($__LOW_STATUS_SUCCESS, 4, $iLong) ; Long from CMYK
		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0) ; wrong number of Parameters
	EndSwitch
EndFunc   ;==>_LOWriter_ConvertColorToLong

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ConvertFromMicrometer
; Description ...: Convert from Micrometer to Inch, Centimeter, Millimeter, or Printer's Points.
; Syntax ........: _LOWriter_ConvertFromMicrometer([$nInchOut = Null[, $nCentimeterOut = Null[, $nMillimeterOut = Null[, $nPointsOut = Null]]]])
; Parameters ....: $nInchOut            - [optional] a general number value. Default is Null. The Micrometers to convert to Inches. See remarks.
;                  $nCentimeterOut      - [optional] a general number value. Default is Null. The Micrometers to convert to Centimeters. See remarks.
;                  $nMillimeterOut      - [optional] a general number value. Default is Null. The Micrometers to convert to Millimeters. See remarks.
;                  $nPointsOut          - [optional] a general number value. Default is Null. The Micrometers to convert to Printer's Points. See remarks.
; Return values .: Success: Number
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $nInchOut not set to Null and not a number.
;				   @Error 1 @Extended 2 Return 0 = $nCentimeterOut not set to Null and not a number.
;				   @Error 1 @Extended 3 Return 0 = $nMillimeterOut not set to Null and not a number.
;				   @Error 1 @Extended 4 Return 0 = $nPointsOut not set to Null and not a number.
;				   @Error 1 @Extended 5 Return 0 = No parameters set to other than Null.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error converting from Micrometers to Inch.
;				   @Error 3 @Extended 2 Return 0 = Error converting from Micrometers to Centimeter.
;				   @Error 3 @Extended 3 Return 0 = Error converting from Micrometers to Millimeter.
;				   @Error 3 @Extended 4 Return 0 = Error converting from Micrometers to Printer's Points.
;				   --Success--
;				   @Error 0 @Extended 1 Return Number. Converted from Micrometers to Inch.
;				   @Error 0 @Extended 2 Return Number. Converted from Micrometers to Centimeter.
;				   @Error 0 @Extended 3 Return Number. Converted from Micrometers to Millimeter.
;				   @Error 0 @Extended 4 Return Number. Converted from Micrometers to Printer's Points.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To skip a parameter, set it to Null. If you are converting to Inches, place the Micrometers in $nInchOut, if
;					converting to Millimeters, $nInchOut and $nCentimeter are set to Null, and $nCMillimetersOut is set.  A
;					Micrometer is 1000th of a centimeter, and is used in almost all Libre Office functions that contain a
;					measurement parameter.
; Related .......: _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ConvertFromMicrometer($nInchOut = Null, $nCentimeterOut = Null, $nMillimeterOut = Null, $nPointsOut = Null)
	Local $nReturnValue

	If ($nInchOut <> Null) Then
		If Not IsNumber($nInchOut) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
		$nReturnValue = __LOWriter_UnitConvert($nInchOut, $__LOWCONST_CONVERT_UM_INCH)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterOut <> Null) Then
		If Not IsNumber($nCentimeterOut) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$nReturnValue = __LOWriter_UnitConvert($nCentimeterOut, $__LOWCONST_CONVERT_UM_CM)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterOut <> Null) Then
		If Not IsNumber($nMillimeterOut) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$nReturnValue = __LOWriter_UnitConvert($nMillimeterOut, $__LOWCONST_CONVERT_UM_MM)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 3, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsOut <> Null) Then
		If Not IsNumber($nPointsOut) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$nReturnValue = __LOWriter_UnitConvert($nPointsOut, $__LOWCONST_CONVERT_UM_PT)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 4, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.
EndFunc   ;==>_LOWriter_ConvertFromMicrometer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ConvertToMicrometer
; Description ...: Convert from Inch, Centimeter, Millimeter, or Printer's Points to Micrometer.
; Syntax ........: _LOWriter_ConvertToMicrometer([$nInchIn = Null[, $nCentimeterIn = Null[, $nMillimeterIn = Null[, $nPointsIn = Null]]]])
; Parameters ....: $nInchIn             - [optional] a general number value. Default is Null. The Inches to convert to Micrometers. See remarks.
;                  $nCentimeterIn       - [optional] a general number value. Default is Null. The Centimeters to convert to Micrometers. See remarks.
;                  $nMillimeterIn       - [optional] a general number value. Default is Null. The Millimeters to convert to Micrometers. See remarks.
;                  $nPointsIn           - [optional] a general number value. Default is Null. The Printer's Points to convert to Micrometers. See remarks.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $nInchIn not set to Null and not a number.
;				   @Error 1 @Extended 2 Return 0 = $nCentimeterIn not set to Null and not a number.
;				   @Error 1 @Extended 3 Return 0 = $nMillimeterIn not set to Null and not a number.
;				   @Error 1 @Extended 4 Return 0 = $nPointsIn not set to Null and not a number.
;				   @Error 1 @Extended 5 Return 0 = No parameters set to other than Null.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error converting from Inch to Micrometers.
;				   @Error 3 @Extended 2 Return 0 = Error converting from Centimeter to Micrometers.
;				   @Error 3 @Extended 3 Return 0 = Error converting from Millimeter to Micrometers.
;				   @Error 3 @Extended 4 Return 0 = Error converting from Printer's Points to Micrometers.
;				   --Success--
;				   @Error 0 @Extended 1 Return Integer. Converted Inches to Micrometers.
;				   @Error 0 @Extended 2 Return Integer. Converted Centimeters to Micrometers.
;				   @Error 0 @Extended 3 Return Integer. Converted Millimeters to Micrometers.
;				   @Error 0 @Extended 4 Return Integer. Converted Printer's Points to Micrometers.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To skip a parameter, set it to Null. If you are converting from Inches, place the inches in $nInchIn, if
;					converting from Centimeters, $nInchIn is set to Null, and $nCentimeters is set. A Micrometer is 1000th of a
;					centimeter, and is used in almost all Libre Office functions that contain a measurement parameter.
; Related .......: _LOWriter_ConvertFromMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ConvertToMicrometer($nInchIn = Null, $nCentimeterIn = Null, $nMillimeterIn = Null, $nPointsIn = Null)
	Local $nReturnValue

	If ($nInchIn <> Null) Then
		If Not IsNumber($nInchIn) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
		$nReturnValue = __LOWriter_UnitConvert($nInchIn, $__LOWCONST_CONVERT_INCH_UM)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $nReturnValue)
	EndIf

	If ($nCentimeterIn <> Null) Then
		If Not IsNumber($nCentimeterIn) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$nReturnValue = __LOWriter_UnitConvert($nCentimeterIn, $__LOWCONST_CONVERT_CM_UM)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 2, $nReturnValue)
	EndIf

	If ($nMillimeterIn <> Null) Then
		If Not IsNumber($nMillimeterIn) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$nReturnValue = __LOWriter_UnitConvert($nMillimeterIn, $__LOWCONST_CONVERT_MM_UM)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 3, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 3, $nReturnValue)
	EndIf

	If ($nPointsIn <> Null) Then
		If Not IsNumber($nPointsIn) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$nReturnValue = __LOWriter_UnitConvert($nPointsIn, $__LOWCONST_CONVERT_PT_UM)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 4, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 4, $nReturnValue)
	EndIf

	Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0) ; NO Unit set.

EndFunc   ;==>_LOWriter_ConvertToMicrometer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyCreate
; Description ...: Create a Date/Time Format Key.
; Syntax ........: _LOWriter_DateFormatKeyCreate(Byref $oDoc, $sFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFormat             - a string value. The Date/Time format String to create.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFormat not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to Create or Retrieve the Format key, but failed.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Format Key was successfully created, returning Format Key integer.
;				   @Error 0 @Extended 1 Return Integer = Success. Format Key already existed, returning Format Key integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_DateFormatKeyDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyCreate(ByRef $oDoc, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $iFormatKey) ; Format already existed
	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $iFormatKey) ; Format created

	Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; Failed to create or retrieve Format
EndFunc   ;==>_LOWriter_DateFormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyDelete
; Description ...: Delete a User-Created Date/Time Format Key from a Document.
; Syntax ........: _LOWriter_DateFormatKeyDelete(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iFormatKey          - an integer value. The User-Created Date/Time format Key to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = Format Key called in $iFormatKey not found in Document.
;				   @Error 1 @Extended 4 Return 0 = Format Key called in $iFormatKey not User-Created.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete key, but Key is still found in Document.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Format Key was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateFormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyDelete(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_DateFormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Key not found.
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0) ; Key not User Created.

	$oFormats.removeByKey($iFormatKey)

	Return (_LOWriter_DateFormatKeyExists($oDoc, $iFormatKey) = False) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_DateFormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyExists
; Description ...: Check if a Document contains a Date/Time Format Key Already or not.
; Syntax ........: _LOWriter_DateFormatKeyExists(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iFormatKey          - an integer value. The Date Format Key to check for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatType Parameter for internal Function not an Integer. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Date/Time Formats.
;				   --Success--
;				   @Error 0 @Extended 0 Return True = Success. Date/Time Format already exists in document.
;				   @Error 0 @Extended 0 Return False = Success. Date/Time Format does not exist in document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyExists(ByRef $oDoc, $iFormatKey)
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = _LOWriter_FormatKeyExists($oDoc, $iFormatKey, $LOW_FORMAT_KEYS_DATE_TIME)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_DateFormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyGetString
; Description ...: Retrieve a Date/Time Format Key String.
; Syntax ........: _LOWriter_DateFormatKeyGetString(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iFormatKey          - an integer value. The Date/Time Format Key to retrieve the string for.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatKey not found in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve requested Format Key Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returning Format Key's Format String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:_LOWriter_DateFormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyGetString(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormatKey

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_DateFormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oFormatKey = $oDoc.getNumberFormats().getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0) ; Failed to retrieve Key

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFormatKey.FormatString())
EndFunc   ;==>_LOWriter_DateFormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateFormatKeyList
; Description ...: Retrieve an Array of Date/Time Format Keys.
; Syntax ........: _LOWriter_DateFormatKeyList(Byref $oDoc[, $bIsUser = False[, $bUserOnly = False[, $bDateOnly = False[, $bTimeOnly = False]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bIsUser             - [optional] a boolean value. Default is False. If True, Adds a third column to the return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Date/Time Format Keys are returned.
;                  $bDateOnly           - [optional] a boolean value. Default is False. If True, Only Date  FormatKeys are returned.
;                  $bTimeOnly           - [optional] a boolean value. Default is False. If True, Only Time Format Keys are returned.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsUser not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bUserOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bDateOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bTimeOnly not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Both $bDateOnly and $bTimeOnly set to True. Set one or both to false.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Date/Time Formats.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning a 2 or three column Array, depending on current $bIsUser setting.
;				   +			Column One (Array[0][0]) will contain the Format Key integer,
;				   +			Column two (Array[0][1]) will contain the Format String
;				   +			And if $bIsUser is set to True, Column Three (Array[0][2]) will contain a Boolean, True if the Format Key is User creater, else false.
;				   +			@Extended is set to the number of Keys returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyDelete, _LOWriter_DateFormatKeyGetString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateFormatKeyList(ByRef $oDoc, $bIsUser = False, $bUserOnly = False, $bDateOnly = False, $bTimeOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avDTFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0, $iQueryType = $LOW_FORMAT_KEYS_DATE_TIME

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDateOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bTimeOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If ($bDateOnly = True) And ($bTimeOnly = True) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)

	$iColumns = ($bIsUser = True) ? $iColumns : 2

	$iQueryType = ($bDateOnly = True) ? $LOW_FORMAT_KEYS_DATE : $iQueryType
	$iQueryType = ($bTimeOnly = True) ? $LOW_FORMAT_KEYS_TIME : $iQueryType

	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$aiFormatKeys = $oFormats.queryKeys($iQueryType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	ReDim $avDTFormats[UBound($aiFormatKeys)][$iColumns]

	For $i = 0 To UBound($aiFormatKeys) - 1

		If ($bUserOnly = True) Then
			If ($oFormats.getbykey($aiFormatKeys[$i]).UserDefined() = True) Then
				$avDTFormats[$iCount][0] = $aiFormatKeys[$i]
				$avDTFormats[$iCount][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
				If ($bIsUser = True) Then $avDTFormats[$iCount][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
				$iCount += 1
			EndIf
		Else
			$avDTFormats[$i][0] = $aiFormatKeys[$i]
			$avDTFormats[$i][1] = $oFormats.getbykey($aiFormatKeys[$i]).FormatString()
			If ($bIsUser = True) Then $avDTFormats[$i][2] = $oFormats.getbykey($aiFormatKeys[$i]).UserDefined()
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	Next

	If ($bUserOnly = True) Then ReDim $avDTFormats[$iCount][$iColumns]

	Return SetError($__LOW_STATUS_SUCCESS, UBound($avDTFormats), $avDTFormats)
EndFunc   ;==>_LOWriter_DateFormatKeyList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateStructCreate
; Description ...: Create a Date Structure for inserting a Date into certain other functions.
; Syntax ........: _LOWriter_DateStructCreate([$iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $iYear               - [optional] an integer value. Default is Null. The Year, in 4 digit integer format.
;                  $iMonth              - [optional] an integer value. Default is Null. The Month, in 2 digit integer format. Set to 0 for Void date. Min 0, Max 12.
;                  $iDay                - [optional] an integer value. Default is Null. The Day, in 2 digit integer format. Set to 0 for Void date. Min 0, Max 31.
;                  $iHours              - [optional] an integer value. Default is Null. The Hour, in 2 digit integer format. Min 0, Max 23.
;                  $iMinutes            - [optional] an integer value. Default is Null. Minutes, in 2 digit integer format. Min 0, Max 59.
;                  $iSeconds            - [optional] an integer value. Default is Null. Seconds, in 2 digit integer format. Min 0, Max 59.
;                  $iNanoSeconds        - [optional] an integer value. Default is Null. Nano-Second, in integer format. Min 0, Max 999,999,999.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If true: time zone is UTC Else false: unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: Structure.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $iYear not an Integer.
;				   @Error 1 @Extended 2 Return 0 = $iYear not 4 digits long.
;				   @Error 1 @Extended 3 Return 0 = $iMonth not an Integer, less than 0 or greater than 12.
;				   @Error 1 @Extended 4 Return 0 = $iDay not an Integer, less than 0 or greater than 31.
;				   @Error 1 @Extended 5 Return 0 = $iHours not an Integer, less than 0 or greater than 23.
;				   @Error 1 @Extended 6 Return 0 = $iMinutes not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 7 Return 0 = $iSeconds not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 8 Return 0 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;				   @Error 1 @Extended 9 Return 0 = $bIsUTC not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.util.DateTime" Object.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;				   --Success--
;				   @Error 0 @Extended 0 Return Structure = Success. Successfully created the Date/Time Structure, Returning the Date/Time Structure Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateStructCreate($iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tDateStruct

	$tDateStruct = __LOWriter_CreateStruct("com.sun.star.util.DateTime")
	If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		$tDateStruct.Year = $iYear
	Else
		$tDateStruct.Year = @YEAR
	EndIf

	If ($iMonth <> Null) Then
		If Not __LOWriter_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tDateStruct.Month = $iMonth
	Else
		$tDateStruct.Month = @MON
	EndIf

	If ($iDay <> Null) Then
		If Not __LOWriter_IntIsBetween($iDay, 0, 31) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tDateStruct.Day = $iDay
	Else
		$tDateStruct.Day = @MDAY
	EndIf

	If ($iHours <> Null) Then
		If Not __LOWriter_IntIsBetween($iHours, 0, 23) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tDateStruct.Hours = $iHours
	Else
		$tDateStruct.Hours = @HOUR
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LOWriter_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tDateStruct.Minutes = $iMinutes
	Else
		$tDateStruct.Minutes = @MIN
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tDateStruct.Seconds = $iSeconds
	Else
		$tDateStruct.Seconds = @SEC
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tDateStruct.NanoSeconds = $iNanoSeconds
	Else
		$tDateStruct.NanoSeconds = 0
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		If Not __LOWriter_VersionCheck(4.1) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$tDateStruct.IsUTC = $bIsUTC
	Else
		If __LOWriter_VersionCheck(4.1) Then $tDateStruct.IsUTC = False
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, $tDateStruct)
EndFunc   ;==>_LOWriter_DateStructCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_DateStructModify
; Description ...: Set or retrieve Date Structure settings.
; Syntax ........: _LOWriter_DateStructModify(Byref $tDateStruct[, $iYear = Null[, $iMonth = Null[, $iDay = Null[, $iHours = Null[, $iMinutes = Null[, $iSeconds = Null[, $iNanoSeconds = Null[, $bIsUTC = Null]]]]]]]])
; Parameters ....: $tDateStruct         - [in/out] a dll struct value. The Date Structure to modify, returned from a _LOWriter_DateStructCreate, or setting retrieval function. Structure will be directly modified.
;                  $iYear               - [optional] an integer value. Default is Null. The Year, in 4 digit integer format.
;                  $iMonth              - [optional] an integer value. Default is Null. The Month, in 2 digit integer format. Set to 0 for Void date. Min 0, Max 12.
;                  $iDay                - [optional] an integer value. Default is Null. The Day, in 2 digit integer format. Set to 0 for Void date. Min 0, Max 31.
;                  $iHours              - [optional] an integer value. Default is Null. The Hour, in 2 digit integer format. Min 0, Max 23.
;                  $iMinutes            - [optional] an integer value. Default is Null. Minutes, in 2 digit integer format. Min 0, Max 59.
;                  $iSeconds            - [optional] an integer value. Default is Null. Seconds, in 2 digit integer format. Min 0, Max 59.
;                  $iNanoSeconds        - [optional] an integer value. Default is Null. Nano-Second, in integer format. Min 0, Max 999,999,999.
;                  $bIsUTC              - [optional] a boolean value. Default is Null. If true: time zone is UTC Else false: unknown time zone. Libre Office version 4.1 and up.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iYear not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iYear not 4 digits long.
;				   @Error 1 @Extended 4 Return 0 = $iMonth not an Integer, less than 0 or greater than 12.
;				   @Error 1 @Extended 5 Return 0 = $iDay not an Integer, less than 0 or greater than 31.
;				   @Error 1 @Extended 6 Return 0 = $iHours not an Integer, less than 0 or greater than 23.
;				   @Error 1 @Extended 7 Return 0 = $iMinutes not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 8 Return 0 = $iSeconds not an Integer, less than 0 or greater than 59.
;				   @Error 1 @Extended 9 Return 0 = $iNanoSeconds not an Integer, less than 0 or greater than 999999999.
;				   @Error 1 @Extended 10 Return 0 = $bIsUTC not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iYear
;				   |								2 = Error setting $iMonth
;				   |								4 = Error setting $iDay
;				   |								8 = Error setting $iHours
;				   |								16 = Error setting $iMinutes
;				   |								32 = Error setting $iSeconds
;				   |								64 = Error setting $iNanoSeconds
;				   |								128 = Error setting $bIsUTC
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.1.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 or 8 Element Array with values in order of function parameters. If current Libre Office version is less than 4.1, the Array will contain 7 elements, as $bIsUTC will be eliminated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_DateStructCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_DateStructModify(ByRef $tDateStruct, $iYear = Null, $iMonth = Null, $iDay = Null, $iHours = Null, $iMinutes = Null, $iSeconds = Null, $iNanoSeconds = Null, $bIsUTC = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avMod[7]

	If Not IsObj($tDateStruct) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iYear, $iMonth, $iDay, $iHours, $iMinutes, $iSeconds, $iNanoSeconds, $bIsUTC) Then
		If __LOWriter_VersionCheck(4.1) Then
			__LOWriter_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds(), $tDateStruct.IsUTC())
		Else
			__LOWriter_ArrayFill($avMod, $tDateStruct.Year(), $tDateStruct.Month(), $tDateStruct.Day(), $tDateStruct.Hours(), _
					$tDateStruct.Minutes(), $tDateStruct.Seconds(), $tDateStruct.NanoSeconds())
		EndIf

		Return SetError($__LOW_STATUS_SUCCESS, 1, $avMod)
	EndIf

	If ($iYear <> Null) Then
		If Not IsInt($iYear) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
		If Not (StringLen($iYear) = 4) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		$tDateStruct.Year = $iYear
		$iError = ($tDateStruct.Year() = $iYear) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iMonth <> Null) Then
		If Not __LOWriter_IntIsBetween($iMonth, 0, 12) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tDateStruct.Month = $iMonth
		$iError = ($tDateStruct.Month() = $iMonth) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iDay <> Null) Then
		If Not __LOWriter_IntIsBetween($iDay, 0, 31) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tDateStruct.Day = $iDay
		$iError = ($tDateStruct.Day() = $iDay) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iHours <> Null) Then
		If Not __LOWriter_IntIsBetween($iHours, 0, 23) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tDateStruct.Hours = $iHours
		$iError = ($tDateStruct.Hours() = $iHours) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iMinutes <> Null) Then
		If Not __LOWriter_IntIsBetween($iMinutes, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tDateStruct.Minutes = $iMinutes
		$iError = ($tDateStruct.Minutes() = $iMinutes) ? $iError : BitOR($iError, 16)
	EndIf

	If ($iSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iSeconds, 0, 59) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tDateStruct.Seconds = $iSeconds
		$iError = ($tDateStruct.Seconds() = $iSeconds) ? $iError : BitOR($iError, 32)
	EndIf

	If ($iNanoSeconds <> Null) Then
		If Not __LOWriter_IntIsBetween($iNanoSeconds, 0, 999999999) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$tDateStruct.NanoSeconds = $iNanoSeconds
		$iError = ($tDateStruct.NanoSeconds() = $iNanoSeconds) ? $iError : BitOR($iError, 64)
	EndIf

	If ($bIsUTC <> Null) Then
		If Not IsBool($bIsUTC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		If Not __LOWriter_VersionCheck(4.1) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$tDateStruct.IsUTC = $bIsUTC
		$iError = ($tDateStruct.IsUTC() = $bIsUTC) ? $iError : BitOR($iError, 128)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_DateStructModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyAlignment
; Description ...: Modify or Add Find Format Alignment Settings.
; Syntax ........: _LOWriter_FindFormatModifyAlignment(Byref $atFormat[, $iHorAlign = Null[, $iVertAlign = Null[, $iLastLineAlign = Null[, $bExpandSingleWord = Null[, $bSnapToGrid = Null[, $iTxtDirection = Null]]]]]])
; Parameters ....: $atFormat            - [in/out] an array of dll structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iHorAlign           - [optional] an integer value (0-3). Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iVertAlign          - [optional] an integer value (0-4). Default is Null. The Vertical alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3. In my personal testing, searching for the Vertical Alignment setting using this parameter causes any results matching the searched for string to be replaced, whether they contain the Vert. Align format or not, this is supposed to be fixed in L.O. 7.6.
;                  $iLastLineAlign      - [optional] an integer value (0-3). Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bExpandSingleWord   - [optional] a boolean value. Default is Null. If the last line of a justified paragraph consists of one word, the word is stretched to the width of the paragraph.
;                  $bSnapToGrid         - [optional] a boolean value. Default is Null. If True, Aligns the paragraph to a text grid (if one is active).
;                  $iTxtDirection       - [optional] an integer value (0-5). Default is Null. The Text Writing Direction. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3. [Libre Office Default is 4] In my personal testing, searching for the Text Direction setting using this parameter alone, without using other parameters, causes any results matching the searched for string to be replaced, whether they contain the Text Direction format or not, this is supposed to be fixed in L.O. 7.6.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iHorAlign not an integer, less than 0 or greater than 3. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iVertAlign not an integer, less than 0 or more than 4. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iLastLineAlign not an integer, less than 0 or more than 3. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $bExpandSingleWord not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $bSnapToGrid not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iTxtDirection not an Integer, less than 0 or greater than 5, See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					Note: $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
; Text Direction Constants: $LOW_TXT_DIR_LR_TB(0), — text within lines is written left-to-right. Lines and blocks are placed top-to-bottom. Typically, this is the writing mode for normal "alphabetic" text.
;							$LOW_TXT_DIR_RL_TB(1), — text within a line are written right-to-left. Lines and blocks are placed top-to-bottom. Typically, this writing mode is used in Arabic and Hebrew text.
;							$LOW_TXT_DIR_TB_RL(2), — text within a line is written top-to-bottom. Lines and blocks are placed right-to-left. Typically, this writing mode is used in Chinese and Japanese text.
;							$LOW_TXT_DIR_TB_LR(3), — text within a line is written top-to-bottom. Lines and blocks are placed left-to-right. Typically, this writing mode is used in Mongolian text.
;							$LOW_TXT_DIR_CONTEXT(4), — obtain actual writing mode from the context of the object.
;							$LOW_TXT_DIR_BT_LR(5), — text within a line is written bottom-to-top. Lines and blocks are placed left-to-right. (LibreOffice 6.3)
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyAlignment(ByRef $atFormat, $iHorAlign = Null, $iVertAlign = Null, $iLastLineAlign = Null, $bExpandSingleWord = Null, $bSnapToGrid = Null, $iTxtDirection = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iHorAlign <> Null) Then
		If ($iHorAlign = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaAdjust")
		Else
			If Not __LOWriter_IntIsBetween($iHorAlign, $LOW_PAR_ALIGN_HOR_LEFT, $LOW_PAR_ALIGN_HOR_CENTER) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaAdjust", $iHorAlign))
		EndIf
	EndIf

	If ($iVertAlign <> Null) Then
		If ($iVertAlign = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaVertAlignment")
		Else
			If Not __LOWriter_IntIsBetween($iVertAlign, $LOW_PAR_ALIGN_VERT_AUTO, $LOW_PAR_ALIGN_VERT_BOTTOM) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaVertAlignment", $iVertAlign))
		EndIf
	EndIf

	If ($iLastLineAlign <> Null) Then
		If ($iLastLineAlign = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaLastLineAdjust")
		Else
			If Not __LOWriter_IntIsBetween($iLastLineAlign, $LOW_PAR_LAST_LINE_JUSTIFIED, $LOW_PAR_LAST_LINE_CENTER, "", $LOW_PAR_LAST_LINE_START) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaLastLineAdjust", $iLastLineAlign))
		EndIf
	EndIf

	If ($bExpandSingleWord <> Null) Then
		If ($bExpandSingleWord = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaExpandSingleWord")
		Else
			If Not IsBool($bExpandSingleWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaExpandSingleWord", $bExpandSingleWord))
		EndIf
	EndIf

	If ($bSnapToGrid <> Null) Then
		If ($bSnapToGrid = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "SnapToGrid")
		Else
			If Not IsBool($bSnapToGrid) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("SnapToGrid", $bSnapToGrid))
		EndIf
	EndIf

	If ($iTxtDirection <> Null) Then
		If ($iTxtDirection = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "WritingMode")
		Else
			If Not __LOWriter_IntIsBetween($iTxtDirection, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("WritingMode", $iTxtDirection))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyEffects
; Description ...: Modify or Add Find Format Effects Settings.
; Syntax ........: _LOWriter_FindFormatModifyEffects(Byref $atFormat[,$iRelief  = Null[, $iCase = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Character Relief style. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3. In my personal testing, searching for the Relief setting using this parameter causes any results matching the searched for string to be replaced, whether they contain the Relief format or not, this is supposed to be fixed in L.O. 7.6.
;                  $iCase               - [optional] an integer value (0-4). Default is Null. The Character Case Style. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3
;                  $bOutline            - [optional] a boolean value. Default is Null. Whether the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. Whether the characters have a shadow.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iRelief not an integer or less than 0 or greater than 2. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iCase not an integer or less than 0 or greater than 4. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $bOutline not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bShadow not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyEffects(ByRef $atFormat, $iRelief = Null, $iCase = Null, $bOutline = Null, $bShadow = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iRelief <> Null) Then
		If ($iRelief = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharRelief")
		Else
			If Not __LOWriter_IntIsBetween($iRelief, $LOW_RELIEF_NONE, $LOW_RELIEF_ENGRAVED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharRelief", $iRelief))
		EndIf
	EndIf

	If ($iCase <> Null) Then
		If ($iCase = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharCaseMap")
		Else
			If Not __LOWriter_IntIsBetween($iCase, $LOW_CASEMAP_NONE, $LOW_CASEMAP_SM_CAPS) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharCaseMap", $iCase))
		EndIf
	EndIf

	If ($bOutline <> Null) Then
		If ($bOutline = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharContoured")
		Else
			If Not IsBool($bOutline) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharContoured", $bOutline))
		EndIf
	EndIf

	If ($bShadow <> Null) Then
		If ($bShadow = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharShadowed")
		Else
			If Not IsBool($bShadow) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharShadowed", $bShadow))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyEffects

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyFont
; Description ...: Modify or Add Find Format Font Settings.
; Syntax ........: _LOWriter_FindFormatModifyFont(Byref $oDoc, Byref $atFormat[, $sFontName = Null[, $iFontSize = Null[, $iFontWeight = Null[, $iFontPosture = Null[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified. See Remarks.
;                  $sFontName           - [optional] a string value. Default is Null. The Font name to search for.
;                  $iFontSize           - [optional] an integer value. Default is Null. The Font size to search for.
;                  $iFontWeight         - [optional] an integer value(0,50-200). Default is Null. The Font weight to search for. See Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iFontPosture        - [optional] an integer value (0-5). Default is Null. The Font Posture(Italic etc.,) See Constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iFontColor          - [optional] an integer value (-1-16777215). Default is Null. The Font Color in Long Integer format, Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iTransparency       - [optional] an integer value. Default is Null. The percentage of Transparency, Min. 0, Max 100. 0 is not visible, 100 is fully visible. Seems to require a color entered in $iFontColor before transparency can be searched for. Libre Office 7.0 and Up.
;                  $iHighlight          - [optional] an integer value (-1-16777215). Default is Null. The Highlight color to search for, in Long Integer format, Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 3 Return 0 = $sFontName not a String.
;				   @Error 1 @Extended 4 Return 0 = Font defined in $sFontName not found in current Document.
;				   @Error 1 @Extended 5 Return 0 = $iFontSize not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iFontWeight not an Integer, less than 50 but not 0, or more than 200. See Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $iFontPosture not an Integer, less than 0 or greater than 5. See Constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;				   @Error 1 @Extended 9 Return 0 = $iTransparency not an Integer, Less than 0 or greater than 100.
;				   @Error 1 @Extended 10 Return 0 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 7.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,_LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll _LOWriter_DocReplaceAllInRange,
;					_LOWriter_FontsList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyFont(ByRef $oDoc, ByRef $atFormat, $sFontName = Null, $iFontSize = Null, $iFontWeight = Null, $iFontPosture = Null, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($sFontName <> Null) Then
		If ($sFontName = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharFontName")
		Else
			If Not IsString($sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			If Not _LOWriter_FontExists($oDoc, $sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharFontName", $sFontName))
		EndIf
	EndIf

	If ($iFontSize <> Null) Then
		If ($iFontSize = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharHeight")
		Else
			If Not IsInt($iFontSize) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharHeight", $iFontSize))
		EndIf
	EndIf

	If ($iFontWeight <> Null) Then
		If ($iFontWeight = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWeight")
		Else
			If Not __LOWriter_IntIsBetween($iFontWeight, $LOW_WEIGHT_THIN, $LOW_WEIGHT_BLACK, "", $LOW_WEIGHT_DONT_KNOW) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWeight", $iFontWeight))
		EndIf
	EndIf

	If ($iFontPosture <> Null) Then
		If ($iFontPosture = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharPosture")
		Else
			If Not __LOWriter_IntIsBetween($iFontPosture, $LOW_POSTURE_NONE, $LOW_POSTURE_ITALIC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharPosture", $iFontPosture))
		EndIf
	EndIf

	If ($iFontColor <> Null) Then
		If ($iFontColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharColor")
		Else
			If Not __LOWriter_IntIsBetween($iFontColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharColor", $iFontColor))
		EndIf
	EndIf

	If ($iTransparency <> Null) Then
		If ($iTransparency = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharTransparence")
		Else
			If Not __LOWriter_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
			If Not __LOWriter_VersionCheck(7.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharTransparence", $iTransparency))
		EndIf
	EndIf

	If ($iHighlight <> Null) Then
		If ($iHighlight = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharBackColor")
			If __LOWriter_VersionCheck(4.2) Then __LOWriter_FindFormatDeleteSetting($atFormat, "CharHighlight")
		Else
			If Not __LOWriter_IntIsBetween($iHighlight, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
			; CharHighlight; same as CharBackColor---Libre seems to use back color for highlighting.
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharBackColor", $iHighlight))
			If __LOWriter_VersionCheck(4.2) Then __LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharHighlight", $iHighlight))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyHyphenation
; Description ...: Modify or Add Find Format Hyphenation Settings. See Remarks.
; Syntax ........: _LOWriter_FindFormatModifyHyphenation(Byref $atFormat[, $bAutoHyphen = Null[, $bHyphenNoCaps = Null[, $iMaxHyphens = Null[, $iMinLeadingChar = Null[, $iMinTrailingChar = Null]]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $bAutoHyphen         - [optional] a boolean value. Default is Null. Whether  automatic hyphenation is applied.
;                  $bHyphenNoCaps       - [optional] a boolean value. Default is Null. Setting to true will disable hyphenation of words written in CAPS for this paragraph. Libre 6.4 and up.
;                  $iMaxHyphens         - [optional] an integer value. Default is Null. The maximum number of consecutive hyphens. Min 0, Max 99.
;                  $iMinLeadingChar     - [optional] an integer value. Default is Null. Specifies the minimum number of characters to remain before the hyphen character (when hyphenation is applied). Min 2, max 9.
;                  $iMinTrailingChar    - [optional] an integer value. Default is Null. Specifies the minimum number of characters to remain after the hyphen character (when hyphenation is applied). Min 2, max 9.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bAutoHyphen not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bHyphenNoCaps not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iMaxHyphens not an Integer, less than 0, or greater than 99.
;				   @Error 1 @Extended 5 Return 0 = $iMinLeadingChar not an Integer, less than 2 or greater than 9.
;				   @Error 1 @Extended 6 Return 0 = $iMinTrailingChar not an Integer, less than 2 or greater than 9.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 6.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In my personal testing, searching for any of these hyphenation formatting settings causes any results
;						matching the searched for string to be replaced, whether they contain these formatting settings or not,
;						I am unsure why.
;					Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyHyphenation(ByRef $atFormat, $bAutoHyphen = Null, $bHyphenNoCaps = Null, $iMaxHyphens = Null, $iMinLeadingChar = Null, $iMinTrailingChar = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bAutoHyphen <> Null) Then
		If ($bAutoHyphen = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaIsHyphenation")
		Else
			If Not IsBool($bAutoHyphen) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaIsHyphenation", $bAutoHyphen))
		EndIf
	EndIf

	If ($bHyphenNoCaps <> Null) Then
		If ($bHyphenNoCaps = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationNoCaps")
		Else
			If Not IsBool($bHyphenNoCaps) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			If Not __LOWriter_VersionCheck(6.4) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationNoCaps", $bHyphenNoCaps))
		EndIf
	EndIf

	If ($iMaxHyphens <> Null) Then
		If ($iMaxHyphens = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationMaxHyphens")
		Else
			If Not __LOWriter_IntIsBetween($iMaxHyphens, 0, 99) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationMaxHyphens", $iMaxHyphens))
		EndIf
	EndIf

	If ($iMinLeadingChar <> Null) Then
		If ($iMinLeadingChar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationMaxLeadingChars")
		Else
			If Not __LOWriter_IntIsBetween($iMinLeadingChar, 2, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationMaxLeadingChars", $iMinLeadingChar))
		EndIf
	EndIf

	If ($iMinTrailingChar <> Null) Then
		If ($iMinTrailingChar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaHyphenationMaxTrailingChars")
		Else
			If Not __LOWriter_IntIsBetween($iMinTrailingChar, 2, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaHyphenationMaxTrailingChars", $iMinTrailingChar))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyHyphenation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyIndent
; Description ...: Modify or Add Find Format Indent Settings.
; Syntax ........: _LOWriter_FindFormatModifyIndent(Byref $atFormat[, $iBeforeText = Null[, $iAfterText = Null[, $iFirstLine = Null[, $bAutoFirstLine = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iBeforeText         - [optional] an integer value. Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in MicroMeters(uM) Min. -9998989, Max.17094. Both $iBeforeText and $iAfterText must be set to perform a search for either.
;                  $iAfterText          - [optional] an integer value. Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in MicroMeters(uM) Min. -9998989, Max.17094. Both $iBeforeText and $iAfterText must be set to perform a search for either.
;                  $iFirstLine          - [optional] an integer value. Default is Null. Indentation distance of the first line of a paragraph, Set in MicroMeters(uM) Min. -57785, Max.17094. Both $iBeforeText and $iAfterText must be set to perform a search for $iFirstLine.
;                  $bAutoFirstLine      - [optional] a boolean value. Default is Null. Whether the first line should be indented automatically. Both $iBeforeText and $iAfterText must be set to perform a search for $bAutoFirstLine.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iBeforeText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 3 Return 0 = $iAfterText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 4 Return 0 = $iFirstLine not an integer, less than -57785 or more than 17094 uM.
;				   @Error 1 @Extended 5 Return 0 = $bAutoFirstLine not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					Note: $iFirstLine Indent cannot be set if $bAutoFirstLine is set to True.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyIndent(ByRef $atFormat, $iBeforeText = Null, $iAfterText = Null, $iFirstLine = Null, $bAutoFirstLine = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	; Min: -9998989;Max: 17094
	If ($iBeforeText <> Null) Then
		If ($iBeforeText = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaLeftMargin")
		Else
			If Not __LOWriter_IntIsBetween($iBeforeText, -9998989, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaLeftMargin", $iBeforeText))
		EndIf
	EndIf

	If ($iAfterText <> Null) Then
		If ($iAfterText = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaRightMargin")
		Else
			If Not __LOWriter_IntIsBetween($iAfterText, -9998989, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaRightMargin", $iAfterText))
		EndIf
	EndIf

	; mx 17094min;-57785
	If ($iFirstLine <> Null) Then
		If ($iFirstLine = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaFirstLineIndent")
		Else
			If Not __LOWriter_IntIsBetween($iFirstLine, -57785, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaFirstLineIndent", $iFirstLine))
		EndIf
	EndIf

	If ($bAutoFirstLine <> Null) Then
		If ($bAutoFirstLine = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaIsAutoFirstLineIndent")
		Else
			If Not IsBool($bAutoFirstLine) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaIsAutoFirstLineIndent", $bAutoFirstLine))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyOverline
; Description ...: Modify or Add Find Format Overline Settings.
; Syntax ........: _LOWriter_FindFormatModifyOverline(Byref $atFormat[, $iOverLineStyle = Null[, $bWordOnly = Null[, $bOLHasColor = Null[, $iOLColor = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iOverLineStyle      - [optional] an integer value (0-18). Default is Null. The style of the Overline line, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3. See remarks. Overline style must be set before any of the other parameters can be searched for.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined. See remarks.
;                  $bOLHasColor         - [optional] a boolean value. Default is Null. Whether the Overline is colored, must be set to true in order to set the Overline color.
;                  $iOLColor            - [optional] an integer value (-1-16777215). Default is Null. The color of the Overline, set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iOverLineStyle not an Integer, or less than 0 or greater than 18. See Constants $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bOLHasColor not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iOLColor not an Integer, or less than -1 or greater than 16777215.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					$bWordOnly applies to Underline, Overline and Strikeout, regardless of which is set to true, one setting applies to all.
;					Underline Constants are used for Overline line style.
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyOverline(ByRef $atFormat, $iOverLineStyle = Null, $bWordOnly = Null, $bOLHasColor = Null, $iOLColor = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iOverLineStyle <> Null) Then
		If ($iOverLineStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharOverline")
		Else
			If Not __LOWriter_IntIsBetween($iOverLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharOverline", $iOverLineStyle))
		EndIf
	EndIf

	If ($bWordOnly <> Null) Then
		If ($bWordOnly = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWordMode")
		Else
			If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWordMode", $bWordOnly))
		EndIf
	EndIf

	If ($bOLHasColor <> Null) Then
		If ($bOLHasColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharOverlineHasColor")
		Else
			If Not IsBool($bOLHasColor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharOverlineHasColor", $bOLHasColor))
		EndIf
	EndIf

	If ($iOLColor <> Null) Then
		If ($iOLColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharOverlineColor")
		Else
			If Not __LOWriter_IntIsBetween($iOLColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharOverlineColor", $iOLColor))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyOverline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyPageBreak
; Description ...: Modify or Add Find Format Page Break Settings. See Remarks.
; Syntax ........: _LOWriter_FindFormatModifyPageBreak(Byref $oDoc, Byref $atFormat[, $iBreakType = Null[, $sPageStyle = Null[, $iPgNumOffSet = Null]]])
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iBreakType          - [optional] an integer value (0-6). Default is Null. The Page Break Type. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3..
;                  $sPageStyle          - [optional] a string value. Default is Null. Creates a page break before the paragraph it belongs to and assigns the value as the name of the new page style to use.
;                  $iPgNumOffSet        - [optional] an integer value. Default is Null. If a page break property is set at a paragraph, this property contains the new value for the page number.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iBreakType not an integer, less than 0 or greater than 6. See constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $sPageStyle not a String.
;				   @Error 1 @Extended 4 Return 0 = Page Style defined in $sPageStyle not found in current document.
;				   @Error 1 @Extended 5 Return 0 = $iPgNumOffSet not an Integer or less than 0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In my personal testing, searching for a page break was very hit and miss, especially when searching with the
;					"PageStyle" Name parameter, and it never worked for searching for PageNumberOffset.
;					Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyPageBreak(ByRef $oDoc, ByRef $atFormat, $iBreakType = Null, $sPageStyle = Null, $iPgNumOffSet = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iBreakType <> Null) Then
		If ($iBreakType = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "BreakType")
		Else
			If Not __LOWriter_IntIsBetween($iBreakType, $LOW_BREAK_NONE, $LOW_BREAK_PAGE_BOTH) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("BreakType", $iBreakType))
		EndIf
	EndIf

	If ($sPageStyle <> Null) Then
		If ($sPageStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "PageStyleName") ; PageDescName -- Not working?
		Else
			If Not IsString($sPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			If Not _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("PageStyleName", $sPageStyle))
		EndIf
	EndIf

	If ($iPgNumOffSet <> Null) Then
		If ($iPgNumOffSet = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "PageNumberOffset")
		Else
			If Not __LOWriter_IntIsBetween($iPgNumOffSet, 0, $iPgNumOffSet) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("PageNumberOffset", $iPgNumOffSet))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyPageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyPosition
; Description ...: Modify or Add Find Format Position Settings.
; Syntax ........: _LOWriter_FindFormatModifyPosition(Byref $atFormat[, $bAutoSuper = Null[, $iSuperScript = Null[, $bAutoSub = Null[, $iSubScript = Null[, $iRelativeSize = Null]]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $bAutoSuper          - [optional] a boolean value. Default is Null.  Whether to active automatic sizing for SuperScript. Note: $iRelativeSize must be set to be able to search for Super/SubScript settings.
;                  $iSuperScript        - [optional] an integer value. Default is Null. SuperScript percentage value. See Remarks. Note: $iRelativeSize must be set to be able to search for Super/SubScript settings.
;                  $bAutoSub            - [optional] a boolean value. Default is Null. Whether to active automatic sizing for SubScript. Note: $iRelativeSize must be set to be able to search for Super/SubScript settings.
;                  $iSubScript          - [optional] an integer value. Default is Null. SubScript percentage value. See Remarks. Note: $iRelativeSize must be set to be able to search for Super/SubScript settings.
;                  $iRelativeSize       - [optional] an integer value. Default is Null. Percentage relative to current font size, Min. 1, Max 100.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bAutoSuper not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bAutoSub not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iSuperScript not an integer, or less than 0, higher than 100 and Not 14000.
;				   @Error 1 @Extended 5 Return 0 = $iSubScript not an integer, or less than -100, higher than 100 and Not (-)14000.
;				   @Error 1 @Extended 6 Return 0 = $iRelativeSize not an integer, or less than 1, higher than 100.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					0 is the normal $iSubScript or $iSuperScript setting.
;					The way LibreOffice is set up Super/SubScript are set in the same setting, Super is a positive number from
;						1 to 100 (percentage), SubScript is a negative number set to 1 to 100 percentage. For the user's
;						convenience this function accepts both positive and negative numbers for SubScript, if a positive number
;						is called for SubScript, it is automatically set to a negative. Automatic Superscript has a integer
;						value of 14000, Auto SubScript has a integer value of -14000. There is no settable setting of Automatic
;						Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two
;						separate parameters to be able to determine if the user wants to automatically set SuperScript or
;						SubScript. If you set both Auto SuperScript to True and Auto SubScript to True, or $iSuperScript to an
;						integer and $iSubScript to an integer, Subscript will be set as it is the last in the line to be set in
;						this function, and thus will over-write any SuperScript settings.
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyPosition(ByRef $atFormat, $bAutoSuper = Null, $iSuperScript = Null, $bAutoSub = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bAutoSuper <> Null) Then
		If ($bAutoSuper = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not IsBool($bAutoSuper) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			; If $bAutoSuper = True set it to 14000 (automatic superScript) else if $iSuperScript is set, let that overwrite
			;	the current setting, else if subscript is true or set to an integer, it will overwrite the setting. If nothing
			; else set SubScript to 1
			$iSuperScript = ($bAutoSuper) ? 14000 : (IsInt($iSuperScript)) ? $iSuperScript : (IsInt($iSubScript) Or ($bAutoSub = True)) ? $iSuperScript : 1
		EndIf
	EndIf

	If ($bAutoSub <> Null) Then
		If ($bAutoSub = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not IsBool($bAutoSub) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			; If $bAutoSub = True set it to -14000 (automatic SubScript) else if $iSubScript is set, let that overwrite
			;	the current setting, else if superscript is true or set to an integer, it will overwrite the setting.
			$iSubScript = ($bAutoSub) ? -14000 : (IsInt($iSubScript)) ? $iSubScript : (IsInt($iSuperScript)) ? $iSubScript : 1
		EndIf
	EndIf

	If ($iSuperScript <> Null) Then
		If ($iSuperScript = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not __LOWriter_IntIsBetween($iSuperScript, 0, 100, "", 14000) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharEscapement", $iSuperScript))
		EndIf
	EndIf

	If ($iSubScript <> Null) Then
		If ($iSubScript = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapement")
		Else
			If Not __LOWriter_IntIsBetween($iSubScript, -100, 100, "", "-14000:14000") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			$iSubScript = ($iSubScript > 0) ? Int("-" & $iSubScript) : $iSubScript
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharEscapement", $iSubScript))
		EndIf
	EndIf

	If ($iRelativeSize <> Null) Then
		If ($iRelativeSize = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharEscapementHeight")
		Else
			If Not __LOWriter_IntIsBetween($iRelativeSize, 1, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharEscapementHeight", $iRelativeSize))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyRotateScaleSpace
; Description ...: Modify or Add Find Format Rotate, Scale, and Space Settings.
; Syntax ........: _LOWriter_FindFormatModifyRotateScaleSpace(Byref $atFormat[, $iRotation = Null[, $iScaleWidth = Null[, $bAutoKerning = Null[, $nKerning = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iRotation           - [optional] an integer value. Default is Null. Degrees to rotate the text. Accepts only 0, 90, and 270 degrees. In my personal testing, searching for the Rotate setting using this parameter causes any results matching the searched for string to be replaced, whether they contain the Rotate format or not, this is supposed to be fixed in L.O. 7.6.
;                  $iScaleWidth         - [optional] an integer value. Default is Null. The percentage to horizontally stretch or compress the text. Min. 1. Max 100. 100 is normal sizing. In my personal testing, searching for the Scale Width setting using this parameter causes any results matching the searched for string to be replaced, whether they contain the Scale Width format or not, this is supposed to be fixed in L.O. 7.6.
;                  $bAutoKerning        - [optional] a boolean value. Default is Null. True applies a spacing in between certain pairs of characters. False = disabled.
;                  $nKerning            - [optional] a general number value. Default is Null. The kerning value of the characters. Min is -2 Pt. Max is 928.8 Pt. See Remarks. Values are in Printer's Points as set in the Libre Office UI.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iRotation not an Integer or not equal to 0, 90 or 270 degrees.
;				   @Error 1 @Extended 3 Return 0 = $iScaleWidth not an Integer or less than 1 or greater than 100.
;				   @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2 or greater than 928.8 Points.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User
;						Display, however the internal setting is measured in MicroMeters. They will be automatically converted
;						from Points to MicroMeters and back for retrieval of settings.
;						The acceptable values are from -2 Pt to  928.8 Pt. the figures can be directly converted easily,
;						however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative
;						MicroMeters internally from 928.9 up to 1000 Pt (Max setting). For example, 928.8Pt is the last
;						correct value, which equals 32766 uM (MicroMeters), after this LibreOffice reports the following:
;						;928.9 Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258. Attempting to set
;						Libre's kerning value to anything over 32768 uM causes a COM exception, and attempting to set the
;						 kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these
;						reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyRotateScaleSpace(ByRef $atFormat, $iRotation = Null, $iScaleWidth = Null, $bAutoKerning = Null, $nKerning = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iRotation <> Null) Then
		If ($iRotation = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharRotation")
		Else
			If Not __LOWriter_IntIsBetween($iRotation, 0, 0, "", "90:270") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			$iRotation = ($iRotation > 0) ? ($iRotation * 10) : $iRotation ;rotation set in hundredths (90 deg = 900 etc)so times by 10.
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharRotation", $iRotation))
		EndIf
	EndIf

	If ($iScaleWidth <> Null) Then ; can't be less than 1%
		If ($iScaleWidth = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharScaleWidth")
		Else
			If Not __LOWriter_IntIsBetween($iScaleWidth, 1, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharScaleWidth", $iScaleWidth))
		EndIf
	EndIf

	If ($bAutoKerning <> Null) Then
		If ($bAutoKerning = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharAutoKerning")
		Else
			If Not IsBool($bAutoKerning) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharAutoKerning", $bAutoKerning))
		EndIf
	EndIf

	If ($nKerning <> Null) Then
		If ($nKerning = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharKerning")
		Else
			If Not __LOWriter_NumIsBetween($nKerning, -2, 928.8) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			$nKerning = __LOWriter_UnitConvert($nKerning, $__LOWCONST_CONVERT_PT_UM)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharKerning", $nKerning))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyRotateScaleSpace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifySpacing
; Description ...: Modify or Add Find Format Spacing Settings.
; Syntax ........: _LOWriter_FindFormatModifySpacing(Byref $atFormat[, $iAbovePar = Null[, $iBelowPar = Null[, $bAddSpace = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iAbovePar           - [optional] an integer value. Default is Null. The Space above a paragraph, in Micrometers. Min 0 Micrometers (uM) Max 10,008 uM.
;                  $iBelowPar           - [optional] an integer value. Default is Null. The Space Below a paragraph, in Micrometers. Min 0, Max 10,008 Micrometers (uM).
;                  $bAddSpace           - [optional] a boolean value. Default is Null. If true, the top and bottom margins of the paragraph should not be applied when the previous and next paragraphs have the same style. Libre Office version 3.6 and up.
;                  $iLineSpcMode        - [optional] an integer value (0-3). Default is Null. The type of the line spacing of a paragraph. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3, also notice min and max values for each. Must set both $iLineSpcMode and $iLineSpcHeight to be able to search either.
;                  $iLineSpcHeight      - [optional] an integer value. Default is Null. This value specifies the spacing of the lines. See Remarks for Minimum and Max values. Must set both $iLineSpcMode and $iLineSpcHeight to be able to search either.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iAbovePar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 3 Return 0 = $iBelowPar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 4 Return 0 = $bAddSpace not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iLineSpcMode Not an integer, less than 0 or greater than 3. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3..
;				   @Error 1 @Extended 6 Return 0 = $iLineSpcHeight Not an integer.
;				   @Error 1 @Extended 7 Return 0 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;				   @Error 1 @Extended 8 Return 0 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less than 0 uM or greater than 10008 uM
;				   @Error 1 @Extended 9 Return 0 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 uM or greater than 10008 uM.
;				   --Initialization Errors--
;				   @Error 2 @Extended 2 Return 0 = Error creating LineSpacing Object.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					Note: The settings in Libre Office, (Single,1.15, 1.5, Double,) Use the Proportional mode, and are just
;						varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;					$iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;					Note: $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- 1 MicroMeter once set.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifySpacing(ByRef $atFormat, $iAbovePar = Null, $iBelowPar = Null, $bAddSpace = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLine
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iAbovePar <> Null) Then
		If ($iAbovePar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaTopMargin")
		Else
			If Not __LOWriter_IntIsBetween($iAbovePar, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaTopMargin", $iAbovePar))
		EndIf
	EndIf

	If ($iBelowPar <> Null) Then
		If ($iBelowPar = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaBottomMargin")
		Else
			If Not __LOWriter_IntIsBetween($iBelowPar, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaBottomMargin", $iBelowPar))
		EndIf
	EndIf

	If ($bAddSpace <> Null) Then
		If ($bAddSpace = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaContextMargin")
		Else
			If Not IsBool($bAddSpace) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			If Not __LOWriter_VersionCheck(3.6) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaContextMargin", $bAddSpace))
		EndIf
	EndIf

	If ($iLineSpcMode <> Null) Or ($iLineSpcHeight <> Null) Then
		If ($iLineSpcMode = Default) Or ($iLineSpcHeight = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaLineSpacing")
		Else
			$tLine = __LOWriter_FindFormatRetrieveSetting($atFormat, "ParaLineSpacing") ; Retrieve the ParaLineSpacing Property to modify if it exists.
			If (@error = 0) And (@extended = 1) Then $tLine = $tLine.Value() ; If retrieval was successful, obtain the Line Space Structure.
			If Not IsObj($tLine) Then $tLine = __LOWriter_CreateStruct("com.sun.star.style.LineSpacing") ; If retrieval was not successful, then create a new one.
			If Not IsObj($tLine) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

			If ($iLineSpcMode <> Default) And ($iLineSpcMode <> Null) Then
				If Not __LOWriter_IntIsBetween($iLineSpcMode, $LOW_LINE_SPC_MODE_PROP, $LOW_LINE_SPC_MODE_FIX) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
				$tLine.Mode = $iLineSpcMode
			EndIf

			If ($iLineSpcHeight <> Default) And ($iLineSpcHeight <> Null) Then
				If Not IsInt($iLineSpcHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
				Switch $tLine.Mode()
					Case $LOW_LINE_SPC_MODE_PROP ;Proportional
						If Not __LOWriter_IntIsBetween($iLineSpcHeight, 6, 65535) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0) ; Min setting on Proportional is 6%
					Case $LOW_LINE_SPC_MODE_MIN, $LOW_LINE_SPC_MODE_LEADING ;Minimum and Leading Modes
						If Not __LOWriter_IntIsBetween($iLineSpcHeight, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
					Case $LOW_LINE_SPC_MODE_FIX ;Fixed Line Spacing Mode
						If Not __LOWriter_IntIsBetween($iLineSpcHeight, 51, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0) ; Min spacing is 51 when Fixed Mode
				EndSwitch
				$tLine.Height = $iLineSpcHeight
			EndIf

			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaLineSpacing", $tLine))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifySpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyStrikeout
; Description ...: Modify or Add Find Format Strikeout Settings.
; Syntax ........: _LOWriter_FindFormatModifyStrikeout(Byref $atFormat[, $bWordOnly = Null[, $bStrikeOut = Null[, $iStrikeLineStyle = Null]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined. See remarks.
;                  $bStrikeOut          - [optional] a boolean value. Default is Null. True = strikeout, False = no strike out.
;                  $iStrikeLineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout Line Style, see constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3..
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bStrikeOut not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iStrikeLineStyle not an Integer, or less than 0 or greater than 8. See Constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3..
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					$bWordOnly applies to Underline, Overline and Strikeout, regardless of which is set to true, one setting applies to all.
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyStrikeout(ByRef $atFormat, $bWordOnly = Null, $bStrikeOut = Null, $iStrikeLineStyle = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bWordOnly <> Null) Then
		If ($bWordOnly = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWordMode")
		Else
			If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWordMode", $bWordOnly))
		EndIf
	EndIf

	If ($bStrikeOut <> Null) Then
		If ($bStrikeOut = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharCrossedOut")
		Else
			If Not IsBool($bStrikeOut) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharCrossedOut", $bStrikeOut))
		EndIf
	EndIf

	If ($iStrikeLineStyle <> Null) Then
		If ($iStrikeLineStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharStrikeout")
		Else
			If Not __LOWriter_IntIsBetween($iStrikeLineStyle, $LOW_STRIKEOUT_NONE, $LOW_STRIKEOUT_X) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharStrikeout", $iStrikeLineStyle))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyStrikeout

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyTxtFlowOpt
; Description ...: Modify or Add Find Format Text Flow Settings.
; Syntax ........: _LOWriter_FindFormatModifyTxtFlowOpt(Byref $atFormat[, $bParSplit = Null[, $bKeepTogether = Null[, $iParOrphans = Null[, $iParWidows = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $bParSplit           - [optional] a boolean value. Default is Null. FALSE prevents the paragraph from getting split into two pages or columns
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. TRUE prevents page or column breaks between this and the following paragraph.
;                  $iParOrphans         - [optional] an integer value. Default is Null. Specifies the minimum number of lines of the paragraph that have to be at bottom of a page if the paragraph is spread over more than one page. Min is 0 (disabled), and cannot be 1. Max is 9. In my personal testing, searching for the Orphan setting using this parameter causes any results matching the searched for string to be replaced, whether they contain the Orphan format or not, I am unsure why.
;                  $iParWidows          - [optional] an integer value. Default is Null. Specifies the minimum number of lines of the paragraph that have to be at top of a page if the paragraph is spread over more than one page. Min is 0 (disabled), and cannot be 1. Max is 9. In my personal testing, searching for the Widow setting using this parameter causes any results matching the searched for string to be replaced, whether they contain the Widow format or not, I am unsure why.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $bParSplit not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bKeepTogether not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iParOrphans not an Integer, less than 0, equal to 1, or greater than 9.
;				   @Error 1 @Extended 5 Return 0 = $iParWidows not an Integer, less than 0, equal to 1, or greater than 9.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
; Related .......: _LOWriter_DocFindAll, _LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll
;					_LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyTxtFlowOpt(ByRef $atFormat, $bParSplit = Null, $bKeepTogether = Null, $iParOrphans = Null, $iParWidows = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($bParSplit <> Null) Then
		If ($bParSplit = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaSplit")
		Else
			If Not IsBool($bParSplit) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaSplit", $bParSplit))
		EndIf
	EndIf

	If ($bKeepTogether <> Null) Then
		If ($bKeepTogether = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaKeepTogether")
		Else
			If Not IsBool($bKeepTogether) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaKeepTogether", $bKeepTogether))
		EndIf
	EndIf

	If ($iParOrphans <> Null) Then
		If ($iParOrphans = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaOrphans")
		Else
			If Not __LOWriter_IntIsBetween($iParOrphans, 0, 9, 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaOrphans", $iParOrphans))
		EndIf
	EndIf

	If ($iParWidows <> Null) Then
		If ($iParWidows = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "ParaWidows")
		Else
			If Not __LOWriter_IntIsBetween($iParWidows, 0, 9, 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("ParaWidows", $iParWidows))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyTxtFlowOpt

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FindFormatModifyUnderline
; Description ...: Modify or Add Find Format Underline Settings.
; Syntax ........: _LOWriter_FindFormatModifyUnderline(Byref $atFormat[, $iUnderLineStyle = Null[, $bWordOnly = Null[, $bULHasColor = Null[, $iULColor = Null]]]])
; Parameters ....: $atFormat            - [in/out] an array of structs. A Find Format Array of Settings to modify. Array will be directly modified.
;                  $iUnderLineStyle     - [optional] an integer value (0-18). Default is Null. The style of the Underline line, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3. Underline style must be set before any of the other parameters can be searched for.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined. See remarks.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. Whether the underline is colored, must be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value (-1-16777215). Default is Null. The color of the underline, set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.. $LOW_COLOR_OFF(-1) is automatic color mode.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $atFormat not an Array or contains more than 1 column.
;				   @Error 1 @Extended 2 Return 0 = $iUnderLineStyle not an Integer, or less than 0 or greater than 18. See Constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3..
;				   @Error 1 @Extended 3 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bULHasColor not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iULColor not an Integer, or less than -1 or greater than 16777215.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. FindFormat Array of Settings was successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;					Call any parameter you wish to delete from an already existing Find Format Array with the Default Keyword.
;					If you do not have a pre-existing FindFormat Array, create and pass an Array with 0 elements. (Local $aArray[0])
;					$bWordOnly applies to Underline, Overline and Strikeout, regardless of which is set to true, one setting applies to all.
; Related .......:_LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong, _LOWriter_DocFindAll,
;					_LOWriter_DocFindAllInRange, _LOWriter_DocFindNext, _LOWriter_DocReplaceAll, _LOWriter_DocReplaceAllInRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FindFormatModifyUnderline(ByRef $atFormat, $iUnderLineStyle = Null, $bWordOnly = Null, $bULHasColor = Null, $iULColor = Null)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($atFormat) Or (UBound($atFormat, $UBOUND_COLUMNS) > 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	If ($iUnderLineStyle <> Null) Then
		If ($iUnderLineStyle = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharUnderline")
		Else
			If Not __LOWriter_IntIsBetween($iUnderLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharUnderline", $iUnderLineStyle))
		EndIf
	EndIf

	If ($bWordOnly <> Null) Then
		If ($bWordOnly = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharWordMode")
		Else
			If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharWordMode", $bWordOnly))
		EndIf
	EndIf

	If ($bULHasColor <> Null) Then
		If ($bULHasColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharUnderlineHasColor")
		Else
			If Not IsBool($bULHasColor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharUnderlineHasColor", $bULHasColor))
		EndIf
	EndIf

	If ($iULColor <> Null) Then
		If ($iULColor = Default) Then
			__LOWriter_FindFormatDeleteSetting($atFormat, "CharUnderlineColor")
		Else
			If Not __LOWriter_IntIsBetween($iULColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
			__LOWriter_FindFormatAddSetting($atFormat, __LOWriter_SetPropertyValue("CharUnderlineColor", $iULColor))
		EndIf
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FindFormatModifyUnderline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyCreate
; Description ...: Create a Format Key.
; Syntax ........: _LOWriter_FormatKeyCreate(Byref $oDoc, $sFormat)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFormat             - a string value. The format String to create.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFormat not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to Create or Retrieve the Format key, but failed.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Format Key was successfully created, returning Format Key integer.
;				   @Error 0 @Extended 1 Return Integer = Success. Format Key already existed, returning Format Key integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormatKeyDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyCreate(ByRef $oDoc, $sFormat)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFormatKey
	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFormat) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$iFormatKey = $oFormats.queryKey($sFormat, $tLocale, False)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $iFormatKey) ; Format already existed
	$iFormatKey = $oFormats.addNew($sFormat, $tLocale)
	If ($iFormatKey > -1) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $iFormatKey) ; Format created

	Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; Failed to create or retrieve Format
EndFunc   ;==>_LOWriter_FormatKeyCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyDelete
; Description ...: Delete a User-Created Format Key from a Document.
; Syntax ........: _LOWriter_FormatKeyDelete(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iFormatKey          - an integer value. The User-Created format Key to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = Format Key called in $iFormatKey not found in Document.
;				   @Error 1 @Extended 4 Return 0 = Format Key called in $iFormatKey not User-Created.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete key, but Key is still found in Document.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Format Key was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormatKeyList, _LOWriter_FormatKeyCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyDelete(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLocale
	Local $oFormats

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Key not found.
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	If ($oFormats.getbykey($iFormatKey).UserDefined() = False) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0) ; Key not User Created.

	$oFormats.removeByKey($iFormatKey)

	Return (_LOWriter_FormatKeyExists($oDoc, $iFormatKey) = False) ? SetError($__LOW_STATUS_SUCCESS, 0, 1) : SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>_LOWriter_FormatKeyDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyExists
; Description ...:Check if a Document contains a certain Format Key.
; Syntax ........: _LOWriter_FormatKeyExists(Byref $oDoc, $iFormatKey, Const $iFormatType)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iFormatKey          - an integer value. The Format Key to look for.
;                  $iFormatType         - [optional] an integer value (0-8196). Default is $LOW_FORMAT_KEYS_ALL. The Formatk Key type to search in. Values can be BitOr's together. See Constants, $LOW_FORMAT_KEYS_* as defined in LibreOfficeWriter_Constants.au3..
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatType not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Number Formats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Date/Time Formats.
;				   --Success--
;				   @Error 0 @Extended 0 Return True = Success. Format Key exists in document.
;				   @Error 0 @Extended 1 Return False = Success. Format Key does not exist in document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyExists(ByRef $oDoc, $iFormatKey, $iFormatType = $LOW_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys[0]
	Local $tLocale

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iFormatType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$aiFormatKeys = $oFormats.queryKeys($iFormatType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	For $i = 0 To UBound($aiFormatKeys) - 1
		If ($aiFormatKeys[$i] = $iFormatKey) Then Return SetError($__LOW_STATUS_SUCCESS, 0, True) ; Doc does contain format Key
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	Next

	Return SetError($__LOW_STATUS_SUCCESS, 1, False) ; Doc does not contain format Key
EndFunc   ;==>_LOWriter_FormatKeyExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyGetString
; Description ...: Retrieve a Format Key String.
; Syntax ........: _LOWriter_FormatKeyGetString(Byref $oDoc, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iFormatKey          - an integer value. The Format Key to retrieve the string for.
; Return values .:Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormatKey not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFormatKey not found in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve requested Format Key Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returning Format Key's Format String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyGetString(ByRef $oDoc, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormatKey

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$oFormatKey = $oDoc.getNumberFormats().getByKey($iFormatKey)
	If Not IsObj($oFormatKey) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0) ; Key not found.

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oFormatKey.FormatString())
EndFunc   ;==>_LOWriter_FormatKeyGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FormatKeyList
; Description ...: Retrieve an Array of Date/Time Format Keys.
; Syntax ........: _LOWriter_FormatKeyList(Byref $oDoc[, $bIsUser = False[, $bUserOnly = False[, $iFormatKeyType = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bIsUser             - [optional] a boolean value. Default is False. If True, Adds a third column to the return Array with a boolean, whether each Key is user-created or not.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True, only user-created Format Keys are returned.
;                  $iFormatKeyType      - [optional] an integer value (0-8196). Default is $LOW_FORMAT_KEYS_ALL. The Formatk Key type to retrieve a list for. Values can be BitOr's together. See Constants, $LOW_FORMAT_KEYS_* as defined in LibreOfficeWriter_Constants.au3..
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsUser not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bUserOnly not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iFormatKeyType not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.lang.Locale" Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve NumberFormats Object.
;				   @Error 2 @Extended 3 Return 0 = Failed to obtain Array of Format Keys.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning a 2 or three column Array, depending on current $bIsUser setting.
;				   +						Column One (Array[0][0]) will contain the Format Key integer,
;				   +						Column two (Array[0][1]) will contain the Format Key String,
;				   +						If $bIsUser is set to True, Column Three (Array[0][2]) will contain a Boolean, True if the Format Key is User created, else false.
;				   +						@Extended is set to the number of Keys returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FormatKeyDelete, _LOWriter_FormatKeyGetString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FormatKeyList(ByRef $oDoc, $bIsUser = False, $bUserOnly = False, $iFormatKeyType = $LOW_FORMAT_KEYS_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFormats
	Local $aiFormatKeys
	Local $avFormats[0][3]
	Local $tLocale
	Local $iColumns = 3, $iCount = 0

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIsUser) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$iColumns = ($bIsUser = True) ? $iColumns : 2

	If Not IsInt($iFormatKeyType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$tLocale = __LOWriter_CreateStruct("com.sun.star.lang.Locale")
	If Not IsObj($tLocale) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oFormats = $oDoc.getNumberFormats()
	If Not IsObj($oFormats) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
	$aiFormatKeys = $oFormats.queryKeys($iFormatKeyType, $tLocale, False)
	If Not IsArray($aiFormatKeys) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

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
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	Next

	If ($bUserOnly = True) Then ReDim $avFormats[$iCount][$iColumns]

	Return SetError($__LOW_STATUS_SUCCESS, UBound($avFormats), $avFormats)
EndFunc   ;==>_LOWriter_FormatKeyList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_PathConvert
; Description ...: Converts the input path to or from a LibreOffice URL notation path.
; Syntax ........: _LOWriter_PathConvert($sFilePath[, $iReturnMode = $LOW_PATHCONV_AUTO_RETURN])
; Parameters ....: $sFilePath           - a string value. Full path to convert in String format.
;                  $iReturnMode         - [optional] an integer value (0-2). Default is $__g_iAutoReturn. Designates what format of path to return. See Constants, $LOW_PATHCONV_* as defined in LibreOfficeWriter_Constants.au3..
; Return values .: Success: String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sFilePath is not a string
;				   @Error 1 @Extended 2 Return 0 = $iReturnMode Not a Integer, less than 0 or greater than 2, see constants, $LOW_PATHCONV_* as defined in LibreOfficeWriter_Constants.au3..
;				   --Success--
;				   @Error 0 @Extended 1 Return String = Returning converted File Path from Libre Office URL.
;				   @Error 0 @Extended 2 Return String = Returning converted path from File Path to Libre Office URL.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibeOffice URL notation is based on the Internet Standard RFC 1738, which means only [0-9],[a-zA-Z] are
;					allowed in paths, most other characters need to be converted into ISO 8859-1 (ISO Latin) such as is found
;					in internet URL's (spaces become %20). See: StarOfficeTM 6.0 Office SuiteA SunTM ONE Software Offering,
;					Basic Programmer's Guide; Page 74
;					The user generally should not even need this function, as I have endeavored to convert any URLs to the
;						appropriate computer path format and any input computer paths to a Libre Office URL.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_PathConvert($sFilePath, $iReturnMode = $LOW_PATHCONV_AUTO_RETURN)
	Local Const $STR_STRIPLEADING = 1
	Local $asURLReplace[9][2] = [["%", "%25"], [" ", "%20"], ["\", "/"], [";", "%3B"], ["#", "%23"], ["^", "%5E"], ["{", "%7B"], _
			["}", "%7D"], ["`", "%60"]]
	Local $iPathSearch, $iFileSearch, $iPartialPCPath, $iPartialFilePath

	If Not IsString($sFilePath) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iReturnMode, $LOW_PATHCONV_AUTO_RETURN, $LOW_PATHCONV_PCPATH_RETURN) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$sFilePath = StringStripWS($sFilePath, $STR_STRIPLEADING)

	$iPathSearch = StringRegExp($sFilePath, "[A-Z]\:\\") ; Search For a Computer Path, as in C:\ etc.
	$iPartialPCPath = StringInStr($sFilePath, "\") ; Search for partial computer Path containing a backslash.
	$iFileSearch = StringInStr($sFilePath, "file:///", 0, 1, 1, 9) ; Search for a full Libre path, which begins with File:///
	$iPartialFilePath = StringInStr($sFilePath, "/") ; Search For a Partial Libre path containing forward slash

	If ($iReturnMode = $LOW_PATHCONV_AUTO_RETURN) Then

		If ($iPathSearch > 0) Or ($iPartialPCPath > 0) Then ;  if file path contains partial or full PC path, set to convert to Libre URL.
			$iReturnMode = $LOW_PATHCONV_OFFICE_RETURN
		ElseIf ($iFileSearch > 0) Or ($iPartialFilePath > 0) Then ;  if file path contains partial or full Libre URL, set to convert to PC Path.
			$iReturnMode = $LOW_PATHCONV_PCPATH_RETURN
		Else ; If file path contains neither above. convert to Libre URL
			$iReturnMode = $LOW_PATHCONV_OFFICE_RETURN
		EndIf
	EndIf

	Switch $iReturnMode

		Case $LOW_PATHCONV_OFFICE_RETURN
			If $iFileSearch > 0 Then Return SetError($__LOW_STATUS_SUCCESS, 2, $sFilePath)
			If ($iPathSearch > 0) Then $sFilePath = "file:///" & $sFilePath

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][0], $asURLReplace[$i][1])
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
			Next
			Return SetError($__LOW_STATUS_SUCCESS, 2, $sFilePath)

		Case $LOW_PATHCONV_PCPATH_RETURN
			If ($iPathSearch > 0) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $sFilePath)
			If ($iFileSearch > 0) Then $sFilePath = StringReplace($sFilePath, "file:///", Null)

			For $i = 0 To (UBound($asURLReplace) - 1)
				$sFilePath = StringReplace($sFilePath, $asURLReplace[$i][1], $asURLReplace[$i][0])
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
			Next
			Return SetError($__LOW_STATUS_SUCCESS, 1, $sFilePath)

	EndSwitch

EndFunc   ;==>_LOWriter_PathConvert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_VersionGet
; Description ...: Retrieve the current Office version.
; Syntax ........: _LOWriter_VersionGet([$bSimpleVersion = False[, $bReturnName = False]])
; Parameters ....: $bSimpleVersion      - [optional] a boolean value. Default is False. If True, returns a two digit version number, such as "7.3", else returns the complex version number, such as "7.3.2.4".
;                  $bReturnName         - [optional] a boolean value. Default is True. If True returns the Program Name, such as "LibreOffice", appended before the version, "LibreOffice 7.3".
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $bSimpleVersion not a Boolean.
;				   @Error 1 @Extended 2 Return 0 = $bReturnName not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.configuration.ConfigurationProvider" Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error setting property value.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returns the Office version in String format.
; Author ........: Laurent Godard as found in Andrew Pitonyak's book; Zizi64 as found on OpenOffice forum.
; Modified ......: donnyh13, modified for Autoit ccompatibility and error checking.
; Remarks .......: From Macro code by Zizi64 found at:
;					https://forum.openoffice.org/en/forum/viewtopic.php?t=91542&sid=7f452d65e58ac1cd3cc6063350b5ada0
;					And Andrew Pitonyak in "Useful Macro Information For OpenOffice.org" Pages 49, 50.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_VersionGet($bSimpleVersion = False, $bReturnName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sAccess = "com.sun.star.configuration.ConfigurationAccess", $sVersionName, $sVersion, $sReturn
	Local $oSettings, $oConfigProvider
	Local $aParamArray[1]

	If Not IsBool($bSimpleVersion) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReturnName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	Local $oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$oConfigProvider = $oServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
	If Not IsObj($oConfigProvider) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$aParamArray[0] = __LOWriter_SetPropertyValue("nodepath", "/org.openoffice.Setup/Product")
	If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	$oSettings = $oConfigProvider.createInstanceWithArguments($sAccess, $aParamArray)

	$sVersionName = $oSettings.getByName("ooName")

	$sVersion = ($bSimpleVersion) ? $oSettings.getByName("ooSetupVersion") : $oSettings.getByName("ooSetupVersionAboutBox")

	$sReturn = ($bReturnName) ? ($sVersionName & " " & $sVersion) : $sVersion

	Return SetError($__LOW_STATUS_SUCCESS, 0, $sReturn)
EndFunc   ;==>_LOWriter_VersionGet
