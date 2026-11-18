#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Math
#include "LibreOfficeMath_Internal.au3"
#include "LibreOfficeMath_Constants.au3"

; Other includes for Math

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Modifying, Setting, etc. L.O. Math Formulas and settings.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOMath_MathFormatAlignment
; _LOMath_MathFormatFonts
; _LOMath_MathFormatFontSizes
; _LOMath_MathFormatSpacing
; _LOMath_MathFormatSpacingBorders
; _LOMath_MathFormatSpacingBrackets
; _LOMath_MathFormatSpacingFractionBars
; _LOMath_MathFormatSpacingFractions
; _LOMath_MathFormatSpacingIndexes
; _LOMath_MathFormatSpacingLimits
; _LOMath_MathFormatSpacingMatrices
; _LOMath_MathFormatSpacingOperators
; _LOMath_MathFormatSpacingSymbols
; _LOMath_MathFormula
; _LOMath_MathTextMode
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatAlignment
; Description ...: Set or Retrieve the current Formula Alignment.
; Syntax ........: _LOMath_MathFormatAlignment(ByRef $oDoc[, $iAlignment = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iAlignment          - [optional] (0-2) Default is Null. The Horizontal alignment of the formula. See Constants, $LOM_FORMULA_ALIGN_HORI_* as defined in LibreOfficeMath_Constants.au3.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Alignment value. See Constants, $LOM_FORMULA_ALIGN_HORI_* as defined in LibreOfficeMath_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iAlignment not an Integer, less than 0 or greater than 2. See Constants, $LOM_FORMULA_ALIGN_HORI_* as defined in LibreOfficeMath_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Alignment value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iAlignment
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing, _LOMath_MathFormula
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatAlignment(ByRef $oDoc, $iAlignment = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iCurrAlignment

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iAlignment) Then
		$iCurrAlignment = $oDoc.Alignment()
		If Not IsInt($iCurrAlignment) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurrAlignment)
	EndIf

	If Not __LO_IntIsBetween($iAlignment, $LOM_FORMULA_ALIGN_HORI_LEFT, $LOM_FORMULA_ALIGN_HORI_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.Alignment = $iAlignment
	$iError = ($oDoc.Alignment() = $iAlignment) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatFonts
; Description ...: Set or Retrieve the current Formula Font settings for the several Font kinds.
; Syntax ........: _LOMath_MathFormatFonts(ByRef $oDoc, $iFontKind[, $sFontName = Null[, $bBold = Null[, $bItalic = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iFontKind           - (0-7) The Font kind to modify. See Constants, $LOM_FMT_FONT_KIND_* as defined in LibreOfficeMath_Constants.au3.
;                  $sFontName           - [optional] Default is Null. The Font name.
;                  $bBold               - [optional] Default is Null. If True, the Font is Bold.
;                  $bItalic             - [optional] Default is Null. If True, the Font is Italicized.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iFontKind not an Integer, less than 0 or greater than 7. See Constants, $LOM_FMT_FONT_KIND_* as defined in LibreOfficeMath_Constants.au3.
;                  @Error: 1, @Extended: 3 = $sFontName not a String.
;                  @Error: 1, @Extended: 4 = Font called in $sFontName not found.
;                  @Error: 1, @Extended: 5 = $bBold not a Boolean.
;                  @Error: 1, @Extended: 6 = $bItalic not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current font name.
;                  @Error: 3, @Extended: 2 = Failed to retrieve current bold state.
;                  @Error: 3, @Extended: 3 = Failed to retrieve current italic state.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $bBold
;                  |                               4 = Error setting $bItalic not a Boolean.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatFontSizes, _LOMath_FontsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatFonts(ByRef $oDoc, $iFontKind, $sFontName = Null, $bBold = Null, $bItalic = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurrFontName
	Local $bIsBold, $bIsItalic
	Local $avFont[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iFontKind, $LOM_FMT_FONT_KIND_MATH, $LOM_FMT_FONT_KIND_FIXED_WIDTH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($sFontName <> Null) And Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sFontName <> Null) And ($sFontName <> "") And Not _LOMath_FontExists($sFontName, $oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($bBold <> Null) And Not IsBool($bBold) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($bItalic <> Null) And Not IsBool($bItalic) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	If __LO_VarsAreNull($sFontName, $bBold, $bItalic) Then
		Switch $iFontKind
			Case $LOM_FMT_FONT_KIND_MATH
				$sCurrFontName = $oDoc.FontNameMath()
				$bIsBold = $oDoc.FontMathIsBold()
				$bIsItalic = $oDoc.FontMathIsItalic()

			Case $LOM_FMT_FONT_KIND_VARIABLES
				$sCurrFontName = $oDoc.FontNameVariables()
				$bIsBold = $oDoc.FontVariablesIsBold()
				$bIsItalic = $oDoc.FontVariablesIsItalic()

			Case $LOM_FMT_FONT_KIND_FUNCTIONS
				$sCurrFontName = $oDoc.FontNameFunctions()
				$bIsBold = $oDoc.FontFunctionsIsBold()
				$bIsItalic = $oDoc.FontFunctionsIsItalic()

			Case $LOM_FMT_FONT_KIND_NUMBERS
				$sCurrFontName = $oDoc.FontNameNumbers()
				$bIsBold = $oDoc.FontNumbersIsBold()
				$bIsItalic = $oDoc.FontNumbersIsItalic()

			Case $LOM_FMT_FONT_KIND_TEXT
				$sCurrFontName = $oDoc.FontNameText()
				$bIsBold = $oDoc.FontTextIsBold()
				$bIsItalic = $oDoc.FontTextIsItalic()

			Case $LOM_FMT_FONT_KIND_SERIF
				$sCurrFontName = $oDoc.CustomFontNameSerif()
				$bIsBold = $oDoc.FontSerifIsBold()
				$bIsItalic = $oDoc.FontSerifIsItalic()

			Case $LOM_FMT_FONT_KIND_SANS_SERIF
				$sCurrFontName = $oDoc.CustomFontNameSans()
				$bIsBold = $oDoc.FontSansIsBold()
				$bIsItalic = $oDoc.FontSansIsItalic()

			Case $LOM_FMT_FONT_KIND_FIXED_WIDTH
				$sCurrFontName = $oDoc.CustomFontNameFixed()
				$bIsBold = $oDoc.FontFixedIsBold()
				$bIsItalic = $oDoc.FontFixedIsItalic()
		EndSwitch

		If Not IsString($sCurrFontName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		If Not IsBool($bIsBold) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		If Not IsBool($bIsItalic) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		__LO_ArrayFill($avFont, $sCurrFontName, $bIsBold, $bIsItalic)

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFont)
	EndIf

	Switch $iFontKind
		Case $LOM_FMT_FONT_KIND_MATH
			If ($sFontName <> Null) Then
				$oDoc.FontNameMath = $sFontName
				$iError = ($oDoc.FontNameMath() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.FontNameMath = $oDoc.FontNameMath() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontMathIsBold = $bBold
				$iError = ($oDoc.FontMathIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.FontNameMath = $oDoc.FontNameMath() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontMathIsItalic = $bItalic
				$iError = ($oDoc.FontMathIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf

		Case $LOM_FMT_FONT_KIND_VARIABLES
			If ($sFontName <> Null) Then
				$oDoc.FontNameVariables = $sFontName
				$iError = ($oDoc.FontNameVariables() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.FontNameVariables = $oDoc.FontNameVariables() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontVariablesIsBold = $bBold
				$iError = ($oDoc.FontVariablesIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.FontNameVariables = $oDoc.FontNameVariables() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontVariablesIsItalic = $bItalic
				$iError = ($oDoc.FontVariablesIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf

		Case $LOM_FMT_FONT_KIND_FUNCTIONS
			If ($sFontName <> Null) Then
				$oDoc.FontNameFunctions = $sFontName
				$iError = ($oDoc.FontNameFunctions() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.FontNameFunctions = $oDoc.FontNameFunctions() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontFunctionsIsBold = $bBold
				$iError = ($oDoc.FontFunctionsIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.FontNameFunctions = $oDoc.FontNameFunctions() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontFunctionsIsItalic = $bItalic
				$iError = ($oDoc.FontFunctionsIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf

		Case $LOM_FMT_FONT_KIND_NUMBERS
			If ($sFontName <> Null) Then
				$oDoc.FontNameNumbers = $sFontName
				$iError = ($oDoc.FontNameNumbers() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.FontNameNumbers = $oDoc.FontNameNumbers() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontNumbersIsBold = $bBold
				$iError = ($oDoc.FontNumbersIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.FontNameNumbers = $oDoc.FontNameNumbers() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontNumbersIsItalic = $bItalic
				$iError = ($oDoc.FontNumbersIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf

		Case $LOM_FMT_FONT_KIND_TEXT
			If ($sFontName <> Null) Then
				$oDoc.FontNameText = $sFontName
				$iError = ($oDoc.FontNameText() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.FontNameText = $oDoc.FontNameText() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontTextIsBold = $bBold
				$iError = ($oDoc.FontTextIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.FontNameText = $oDoc.FontNameText() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontTextIsItalic = $bItalic
				$iError = ($oDoc.FontTextIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf

		Case $LOM_FMT_FONT_KIND_SERIF
			If ($sFontName <> Null) Then
				$oDoc.CustomFontNameSerif = $sFontName
				$iError = ($oDoc.CustomFontNameSerif() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.CustomFontNameSerif = $oDoc.CustomFontNameSerif() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontSerifIsBold = $bBold
				$iError = ($oDoc.FontSerifIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.CustomFontNameSerif = $oDoc.CustomFontNameSerif() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontSerifIsItalic = $bItalic
				$iError = ($oDoc.FontSerifIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf

		Case $LOM_FMT_FONT_KIND_SANS_SERIF
			If ($sFontName <> Null) Then
				$oDoc.CustomFontNameSans = $sFontName
				$iError = ($oDoc.CustomFontNameSans() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.CustomFontNameSans = $oDoc.CustomFontNameSans() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontSansIsBold = $bBold
				$iError = ($oDoc.FontSansIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.CustomFontNameSans = $oDoc.CustomFontNameSans() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontSansIsItalic = $bItalic
				$iError = ($oDoc.FontSansIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf

		Case $LOM_FMT_FONT_KIND_FIXED_WIDTH
			If ($sFontName <> Null) Then
				$oDoc.CustomFontNameFixed = $sFontName
				$iError = ($oDoc.CustomFontNameFixed() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bBold <> Null) Then
				$oDoc.CustomFontNameFixed = $oDoc.CustomFontNameFixed() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontFixedIsBold = $bBold
				$iError = ($oDoc.FontFixedIsBold() = $bBold) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bItalic <> Null) Then
				$oDoc.CustomFontNameFixed = $oDoc.CustomFontNameFixed() ; Have to re-apply the font name, otherwise I can't set Bold/italic.
				$oDoc.FontFixedIsItalic = $bItalic
				$iError = ($oDoc.FontFixedIsItalic() = $bItalic) ? ($iError) : (BitOR($iError, 4))
			EndIf
	EndSwitch

	; Have to re-apply font size, otherwise fonts are displayed in puny font.
	$oDoc.BaseFontHeight = $oDoc.BaseFontHeight()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatFonts

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatFontSizes
; Description ...: Set or Retrieve Formula Font sizes.
; Syntax ........: _LOMath_MathFormatFontSizes(ByRef $oDoc[, $iBase = Null[, $iText = Null[, $iIndexes = Null[, $iFunctions = Null[, $iOperators = Null[, $iLimits = Null]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iBase               - [optional] (4-127) Default is Null. The Base Formula Font Size, set in Points.
;                  $iText               - [optional] (5-200) Default is Null. The relative percentage of the base size to make Text in a Formula.
;                  $iIndexes            - [optional] (5-200) Default is Null. The relative percentage of the base size to make Indices in a Formula.
;                  $iFunctions          - [optional] (5-200) Default is Null. The relative percentage of the base size to make Functions in a Formula.
;                  $iOperators          - [optional] (5-200) Default is Null. The relative percentage of the base size to make Operators in a Formula.
;                  $iLimits             - [optional] (5-200) Default is Null. The relative percentage of the base size to make Limits in a Formula.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iBase not an Integer, less than 4 or greater than 127.
;                  @Error: 1, @Extended: 3 = $iText not an Integer, less than 5 or greater than 200.
;                  @Error: 1, @Extended: 4 = $iIndexes not an Integer, less than 5 or greater than 200.
;                  @Error: 1, @Extended: 5 = $iFunctions not an Integer, less than 5 or greater than 200.
;                  @Error: 1, @Extended: 6 = $iOperators not an Integer, less than 5 or greater than 200.
;                  @Error: 1, @Extended: 7 = $iLimits not an Integer, less than 5 or greater than 200.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iBase
;                  |                               2 = Error setting $iText
;                  |                               4 = Error setting $iIndexes
;                  |                               8 = Error setting $iFunctions
;                  |                               16 = Error setting $iOperators
;                  |                               32 = Error setting $iLimits
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatFonts, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatFontSizes(ByRef $oDoc, $iBase = Null, $iText = Null, $iIndexes = Null, $iFunctions = Null, $iOperators = Null, $iLimits = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFontSize[6]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iBase, $iText, $iIndexes, $iFunctions, $iOperators, $iLimits) Then
		__LO_ArrayFill($avFontSize, $oDoc.BaseFontHeight(), $oDoc.RelativeFontHeightText(), $oDoc.RelativeFontHeightIndices(), $oDoc.RelativeFontHeightFunctions(), _
				$oDoc.RelativeFontHeightOperators(), $oDoc.RelativeFontHeightLimits())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFontSize)
	EndIf

	If ($iBase <> Null) Then
		If Not __LO_IntIsBetween($iBase, 4, 127) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.BaseFontHeight = $iBase
		$iError = ($oDoc.BaseFontHeight() = $iBase) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iText <> Null) Then
		If Not __LO_IntIsBetween($iText, 5, 200) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeFontHeightText = $iText
		$iError = ($oDoc.RelativeFontHeightText() = $iText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iIndexes <> Null) Then
		If Not __LO_IntIsBetween($iIndexes, 5, 200) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDoc.RelativeFontHeightIndices = $iIndexes
		$iError = ($oDoc.RelativeFontHeightIndices() = $iIndexes) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iFunctions <> Null) Then
		If Not __LO_IntIsBetween($iFunctions, 5, 200) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDoc.RelativeFontHeightFunctions = $iFunctions
		$iError = ($oDoc.RelativeFontHeightFunctions() = $iFunctions) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iOperators <> Null) Then
		If Not __LO_IntIsBetween($iOperators, 5, 200) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oDoc.RelativeFontHeightOperators = $iOperators
		$iError = ($oDoc.RelativeFontHeightOperators() = $iOperators) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iLimits <> Null) Then
		If Not __LO_IntIsBetween($iLimits, 5, 200) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oDoc.RelativeFontHeightLimits = $iLimits
		$iError = ($oDoc.RelativeFontHeightLimits() = $iLimits) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatFontSizes

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacing
; Description ...: Set or Retrieve Formula Spacing values.
; Syntax ........: _LOMath_MathFormatSpacing(ByRef $oDoc[, $iSpacing = Null[, $iLineSpacing = Null[, $iRootSpacing = Null]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iSpacing            - [optional] (0-200) Default is Null. The percentage of spacing between variables and operators.
;                  $iLineSpacing        - [optional] (0-100) Default is Null. The percentage of spacing between lines.
;                  $iRootSpacing        - [optional] (0-100) Default is Null. The percentage of spacing between the root sign and radicals.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iSpacing not an Integer, less than 0 or greater than 200.
;                  @Error: 1, @Extended: 3 = $iLineSpacing not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 4 = $iRootSpacing not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iSpacing
;                  |                               2 = Error setting $iLineSpacing
;                  |                               4 = Error setting $iRootSpacing
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatAlignment, _LOMath_MathFormatSpacingBorders
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacing(ByRef $oDoc, $iSpacing = Null, $iLineSpacing = Null, $iRootSpacing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iSpacing, $iLineSpacing, $iRootSpacing) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeSpacing(), $oDoc.RelativeLineSpacing(), $oDoc.RelativeRootSpacing())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iSpacing <> Null) Then
		If Not __LO_IntIsBetween($iSpacing, 0, 200) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeSpacing = $iSpacing
		$iError = ($oDoc.RelativeSpacing() = $iSpacing) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iLineSpacing <> Null) Then
		If Not __LO_IntIsBetween($iLineSpacing, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeLineSpacing = $iLineSpacing
		$iError = ($oDoc.RelativeLineSpacing() = $iLineSpacing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iRootSpacing <> Null) Then
		If Not __LO_IntIsBetween($iRootSpacing, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDoc.RelativeRootSpacing = $iRootSpacing
		$iError = ($oDoc.RelativeRootSpacing() = $iRootSpacing) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingBorders
; Description ...: Set or Retrieve Formula Border Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingBorders(ByRef $oDoc[, $iLeft = Null[, $iRight = Null[, $iTop = Null[, $iBottom = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iLeft               - [optional] (0-10000) Default is Null. The space between the Left of the Formula and the edge of the document. Set in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] (0-10000) Default is Null. The space between the Right of the Formula and the edge of the document. Set in Hundredths of a Millimeter (HMM).
;                  $iTop                - [optional] (0-10000) Default is Null. The space between the Top of the Formula and the edge of the document. Set in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] (0-10000) Default is Null. The space between the Bottom of the Formula and the edge of the document. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iLeft not an Integer, less than 0 or greater than 10000.
;                  @Error: 1, @Extended: 3 = $iRight not an Integer, less than 0 or greater than 10000.
;                  @Error: 1, @Extended: 4 = $iTop not an Integer, less than 0 or greater than 10000.
;                  @Error: 1, @Extended: 5 = $iBottom not an Integer, less than 0 or greater than 10000.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLeft
;                  |                               2 = Error setting $iRight
;                  |                               4 = Error setting $iTop
;                  |                               8 = Error setting $iBottom
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingBorders(ByRef $oDoc, $iLeft = Null, $iRight = Null, $iTop = Null, $iBottom = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLeft, $iRight, $iTop, $iBottom) Then
		__LO_ArrayFill($avSpacing, $oDoc.LeftMargin(), $oDoc.RightMargin(), $oDoc.TopMargin(), $oDoc.BottomMargin())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0, 10000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.LeftMargin = $iLeft
		$iError = ($oDoc.LeftMargin() = $iLeft) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0, 10000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RightMargin = $iRight
		$iError = ($oDoc.RightMargin() = $iRight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0, 10000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDoc.TopMargin = $iTop
		$iError = ($oDoc.TopMargin() = $iTop) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0, 10000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDoc.BottomMargin = $iBottom
		$iError = ($oDoc.BottomMargin() = $iBottom) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingBorders

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingBrackets
; Description ...: Set or Retrieve Formula Bracket Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingBrackets(ByRef $oDoc[, $iExcess = Null[, $iSpacing = Null[, $bScaleAll = Null[, $iExcessSize = Null]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iExcess             - [optional] (0-100) Default is Null. The percentage of the vertical distance between the upper edge of the contents and the upper end of the brackets.
;                  $iSpacing            - [optional] (0-100) Default is Null. The percentage of the horizontal distance between the contents and the upper end of the brackets.
;                  $bScaleAll           - [optional] Default is Null. If True, Scales all types of brackets.
;                  $iExcessSize         - [optional] (0-100) Default is Null. The percentage to adjust the excess size.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iExcess not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iSpacing not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 4 = $bScaleAll not a Boolean.
;                  @Error: 1, @Extended: 5 = $iExcessSize not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iExcess
;                  |                               2 = Error setting $iSpacing
;                  |                               4 = Error setting $bScaleAll
;                  |                               8 = Error setting $iExcessSize
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing, _LOMath_MathFormatSpacingSymbols
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingBrackets(ByRef $oDoc, $iExcess = Null, $iSpacing = Null, $bScaleAll = Null, $iExcessSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iExcess, $iSpacing, $bScaleAll, $iExcessSize) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeBracketExcessSize(), $oDoc.RelativeBracketDistance(), $oDoc.IsScaleAllBrackets(), _
				$oDoc.RelativeScaleBracketExcessSize())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iExcess <> Null) Then
		If Not __LO_IntIsBetween($iExcess, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeBracketExcessSize = $iExcess
		$iError = ($oDoc.RelativeBracketExcessSize() = $iExcess) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iSpacing <> Null) Then
		If Not __LO_IntIsBetween($iSpacing, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeBracketDistance = $iSpacing
		$iError = ($oDoc.RelativeBracketDistance() = $iSpacing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bScaleAll <> Null) Then
		If Not IsBool($bScaleAll) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oDoc.IsScaleAllBrackets = $bScaleAll
		$iError = ($oDoc.IsScaleAllBrackets() = $bScaleAll) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iExcessSize <> Null) Then
		If Not __LO_IntIsBetween($iExcessSize, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oDoc.RelativeScaleBracketExcessSize = $iExcessSize
		$iError = ($oDoc.RelativeScaleBracketExcessSize() = $iExcessSize) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingBrackets

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingFractionBars
; Description ...: Set or Retrieve Formula Fraction Bar Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingFractionBars(ByRef $oDoc[, $iExcessLength = Null[, $iWeight = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iExcessLength       - [optional] (0-100) Default is Null. The percentage of the excess length of the fraction line.
;                  $iWeight             - [optional] (1-100) Default is Null. The percentage of the weight of the fraction line.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iExcessLength not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iWeight not an Integer, less than 1 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iExcessLength
;                  |                               2 = Error setting $iWeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing, _LOMath_MathFormatSpacingFractions
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingFractionBars(ByRef $oDoc, $iExcessLength = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iExcessLength, $iWeight) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeFractionBarExcessLength(), $oDoc.RelativeFractionBarLineWeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iExcessLength <> Null) Then
		If Not __LO_IntIsBetween($iExcessLength, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeFractionBarExcessLength = $iExcessLength
		$iError = ($oDoc.RelativeFractionBarExcessLength() = $iExcessLength) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iWeight <> Null) Then
		If Not __LO_IntIsBetween($iWeight, 1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeFractionBarLineWeight = $iWeight
		$iError = ($oDoc.RelativeFractionBarLineWeight() = $iWeight) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingFractionBars

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingFractions
; Description ...: Set or Retrieve Formula Fraction Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingFractions(ByRef $oDoc[, $iNumerator = Null[, $iDenominator = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iNumerator          - [optional] (0-100) Default is Null. The percentage of spacing between the fraction bar and the numerator.
;                  $iDenominator        - [optional] (0-100) Default is Null. The percentage of spacing between the fraction bar and the denominator.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iNumerator not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iDenominator not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iNumerator
;                  |                               2 = Error setting $iDenominator
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing, _LOMath_MathFormatSpacingFractionBars
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingFractions(ByRef $oDoc, $iNumerator = Null, $iDenominator = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iNumerator, $iDenominator) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeFractionNumeratorHeight(), $oDoc.RelativeFractionDenominatorDepth())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iNumerator <> Null) Then
		If Not __LO_IntIsBetween($iNumerator, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeFractionNumeratorHeight = $iNumerator
		$iError = ($oDoc.RelativeFractionNumeratorHeight() = $iNumerator) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iDenominator <> Null) Then
		If Not __LO_IntIsBetween($iDenominator, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeFractionDenominatorDepth = $iDenominator
		$iError = ($oDoc.RelativeFractionDenominatorDepth() = $iDenominator) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingFractions

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingIndexes
; Description ...: Set or Retrieve Formula Index Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingIndexes(ByRef $oDoc[, $iSuperscript = Null[, $iSubscript = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iSuperscript        - [optional] (0-100) Default is Null. The percentage of spacing for superscript indexes.
;                  $iSubscript          - [optional] (0-100) Default is Null. The percentage of spacing for subscript indexes.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iSuperscript not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iSubscript not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iSuperscript
;                  |                               2 = Error setting $iSubscript
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingIndexes(ByRef $oDoc, $iSuperscript = Null, $iSubscript = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iSuperscript, $iSubscript) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeIndexSuperscript(), $oDoc.RelativeIndexSubscript())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iSuperscript <> Null) Then
		If Not __LO_IntIsBetween($iSuperscript, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeIndexSuperscript = $iSuperscript
		$iError = ($oDoc.RelativeIndexSuperscript() = $iSuperscript) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iSubscript <> Null) Then
		If Not __LO_IntIsBetween($iSubscript, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeIndexSubscript = $iSubscript
		$iError = ($oDoc.RelativeIndexSubscript() = $iSubscript) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingIndexes

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingLimits
; Description ...: Set or Retrieve Formula Limits Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingLimits(ByRef $oDoc[, $iUpper = Null[, $iLower = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iUpper              - [optional] (0-100) Default is Null. The percentage of the spacing between the sum symbol and the upper limit.
;                  $iLower              - [optional] (0-100) Default is Null. The percentage of the spacing between the sum symbol and the lower limit.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iUpper not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iLower not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iUpper
;                  |                               2 = Error setting $iLower
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingLimits(ByRef $oDoc, $iUpper = Null, $iLower = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iUpper, $iLower) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeUpperLimitDistance(), $oDoc.RelativeLowerLimitDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iUpper <> Null) Then
		If Not __LO_IntIsBetween($iUpper, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeUpperLimitDistance = $iUpper
		$iError = ($oDoc.RelativeUpperLimitDistance() = $iUpper) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iLower <> Null) Then
		If Not __LO_IntIsBetween($iLower, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeLowerLimitDistance = $iLower
		$iError = ($oDoc.RelativeLowerLimitDistance() = $iLower) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingLimits

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingMatrices
; Description ...: Set or Retrieve Formula Matrix Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingMatrices(ByRef $oDoc[, $iLineSpacing = Null[, $iColumnSpacing = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iLineSpacing        - [optional] (0-300) Default is Null. The percentage of the spacing between matrix elements in a row.
;                  $iColumnSpacing      - [optional] (0-300) Default is Null. The percentage of the spacing between matrix elements in a column.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iLineSpacing not an Integer, less than 0 or greater than 300.
;                  @Error: 1, @Extended: 3 = $iColumnSpacing not an Integer, less than 0 or greater than 300.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLineSpacing
;                  |                               2 = Error setting $iColumnSpacing
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingMatrices(ByRef $oDoc, $iLineSpacing = Null, $iColumnSpacing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLineSpacing, $iColumnSpacing) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeMatrixLineSpacing(), $oDoc.RelativeMatrixColumnSpacing())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iLineSpacing <> Null) Then
		If Not __LO_IntIsBetween($iLineSpacing, 0, 300) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeMatrixLineSpacing = $iLineSpacing
		$iError = ($oDoc.RelativeMatrixLineSpacing() = $iLineSpacing) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iColumnSpacing <> Null) Then
		If Not __LO_IntIsBetween($iColumnSpacing, 0, 300) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeMatrixColumnSpacing = $iColumnSpacing
		$iError = ($oDoc.RelativeMatrixColumnSpacing() = $iColumnSpacing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingMatrices

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingOperators
; Description ...: Set or Retrieve Formula Operator Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingOperators(ByRef $oDoc[, $iExcessSize = Null[, $iSpacing = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iExcessSize         - [optional] Default is Null. The percentage of the height from the variable to the operator's upper edge.
;                  $iSpacing            - [optional] Default is Null. The percentage of the horizontal distance between operators and variables.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iExcessSize not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iSpacing not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iExcessSize
;                  |                               2 = Error setting $iSpacing
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing, _LOMath_MathFormatSpacingSymbols, _LOMath_MathFormatSpacingBrackets
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingOperators(ByRef $oDoc, $iExcessSize = Null, $iSpacing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iExcessSize, $iSpacing) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeOperatorExcessSize(), $oDoc.RelativeOperatorSpacing())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iExcessSize <> Null) Then
		If Not __LO_IntIsBetween($iExcessSize, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeOperatorExcessSize = $iExcessSize
		$iError = ($oDoc.RelativeOperatorExcessSize() = $iExcessSize) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iSpacing <> Null) Then
		If Not __LO_IntIsBetween($iSpacing, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeOperatorSpacing = $iSpacing
		$iError = ($oDoc.RelativeOperatorSpacing() = $iSpacing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingOperators

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormatSpacingSymbols
; Description ...: Set or Retrieve Formula Symbol Spacing values.
; Syntax ........: _LOMath_MathFormatSpacingSymbols(ByRef $oDoc[, $iPrimary = Null[, $iMinimum = Null]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $iPrimary            - [optional] (0-100) Default is Null. The percentage of the height of the symbols in relation to the baseline.
;                  $iMinimum            - [optional] (0-100) Default is Null. The percentage of the minimum distance between a symbol and variable.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $iPrimary not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 3 = $iMinimum not an Integer, less than 0 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iPrimary
;                  |                               2 = Error setting $iMinimum
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatSpacing, _LOMath_MathFormatSpacingOperators, _LOMath_MathFormatSpacingBrackets
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormatSpacingSymbols(ByRef $oDoc, $iPrimary = Null, $iMinimum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avSpacing[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iPrimary, $iMinimum) Then
		__LO_ArrayFill($avSpacing, $oDoc.RelativeSymbolPrimaryHeight(), $oDoc.RelativeSymbolMinimumHeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iPrimary <> Null) Then
		If Not __LO_IntIsBetween($iPrimary, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDoc.RelativeSymbolPrimaryHeight = $iPrimary
		$iError = ($oDoc.RelativeSymbolPrimaryHeight() = $iPrimary) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iMinimum <> Null) Then
		If Not __LO_IntIsBetween($iMinimum, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDoc.RelativeSymbolMinimumHeight = $iMinimum
		$iError = ($oDoc.RelativeSymbolMinimumHeight() = $iMinimum) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormatSpacingSymbols

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathFormula
; Description ...: Set or Retrieve the current Formula.
; Syntax ........: _LOMath_MathFormula(ByRef $oDoc[, $sFormula = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $sFormula            - [optional] Default is Null. The Math Formula.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current Formula as a String
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $sFormula not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Formula.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFormula
; Author ........: donnyh13
; Modified ......:
; Remarks .......: When setting the Formula, the current formula will be overwritten.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LOMath_MathTextMode, _LOMath_MathFormatFonts, _LOMath_MathFormatAlignment
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathFormula(ByRef $oDoc, $sFormula = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $sCurrFormula

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sFormula) Then
		$sCurrFormula = $oDoc.Formula()
		If Not IsString($sCurrFormula) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrFormula)
	EndIf

	If Not IsString($sFormula) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.Formula = $sFormula
	$iError = ($oDoc.Formula() = $sFormula) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathFormula

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOMath_MathTextMode
; Description ...: Set or Retrieve the current Text Mode value.
; Syntax ........: _LOMath_MathTextMode(ByRef $oDoc[, $bTextMode = Null])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOMath_DocOpen, _LOMath_DocConnect, or _LOMath_DocCreate function.
;                  $bTextMode           - [optional] Default is Null. If True, Text Mode is enabled.
; Return values .: Success: 1 or Boolean.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Boolean = Success. All optional parameters were called with Null, returning current TextMode state.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bTextMode not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Alignment value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bTextMode
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOMath_MathFormatAlignment, _LOMath_MathFormatFonts, _LOMath_MathFormula
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOMath_MathTextMode(ByRef $oDoc, $bTextMode = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOMath_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $bCurrTextMode

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bTextMode) Then
		$bCurrTextMode = $oDoc.IsTextMode()
		If Not IsBool($bCurrTextMode) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $bCurrTextMode)
	EndIf

	If Not IsBool($bTextMode) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.IsTextMode = $bTextMode
	$iError = ($oDoc.IsTextMode() = $bTextMode) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOMath_MathTextMode
