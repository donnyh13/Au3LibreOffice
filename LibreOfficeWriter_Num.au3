#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; Other includes for Writer
#include "LibreOfficeWriter_Char.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Applying L.O. Writer Numbering Styles.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_NumStyleCreate
; _LOWriter_NumStyleCurrent
; _LOWriter_NumStyleCustomize
; _LOWriter_NumStyleDelete
; _LOWriter_NumStyleExists
; _LOWriter_NumStyleGetObj
; _LOWriter_NumStyleOrganizer
; _LOWriter_NumStylePosition
; _LOWriter_NumStyleSetLevel
; _LOWriter_NumStylesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleCreate
; Description ...: Create a new Numbering Style in a Document.
; Syntax ........: _LOWriter_NumStyleCreate(ByRef $oDoc, $sNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNumStyle           - a string value. The Name of the new Numbering Style to create.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sNumStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = Numbering Style name called in $sNumStyle already exists in this document.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating "com.sun.star.style.NumberingStyle" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Retrieving "NumberingStyle" Object.
;                  @Error 3 @Extended 2 Return 0 = Error creating new Numbering Style by Name.
;                  @Error 3 @Extended 3 Return 0 = Error Retrieving New Numbering Style Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. New Numbering Style successfully created. Returning Numbering Style Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_NumStyleDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleCreate(ByRef $oDoc, $sNumStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumStyles, $oStyle, $oNumStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oNumStyles = $oDoc.StyleFamilies().getByName("NumberingStyles")
	If Not IsObj($oNumStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If _LOWriter_NumStyleExists($oDoc, $sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oStyle = $oDoc.createInstance("com.sun.star.style.NumberingStyle")
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oNumStyles.insertByName($sNumStyle, $oStyle)

	If Not $oNumStyles.hasByName($sNumStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oNumStyle = $oNumStyles.getByName($sNumStyle)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNumStyle)
EndFunc   ;==>_LOWriter_NumStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleCurrent
; Description ...: Set or Retrieve the current numbering style for a paragraph by Cursor or paragraph Object.
; Syntax ........: _LOWriter_NumStyleCurrent(ByRef $oDoc, ByRef $oObj[, $sNumStyle = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object returned from _LOWriter_ParObjCreateList function.
;                  $sNumStyle           - [optional] a string value. Default is Null. The Numbering Style name to set the paragraph to.
; Return values .: Success: 1 or String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oObj does not support Paragraph Properties Service.
;                  @Error 1 @Extended 4 Return 0 = $sNumStyle not a String.
;                  @Error 1 @Extended 5 Return 0 = Numbering Style called in $sNumStyle doesn't exist in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve current Numbering Style.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sNumStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Numbering Style successfully set.
;                  @Error 0 @Extended 1 Return String = Success. All optional parameters were called with Null, returning current Numbering Style set for the selection.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
; Related .......: _LOWriter_ParObjCreateList, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_NumStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleCurrent(ByRef $oDoc, ByRef $oObj, $sNumStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurrStyle
	Local $iError = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sNumStyle) Then
		$sCurrStyle = $oObj.NumberingStyleName()
		If Not IsString($sCurrStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurrStyle)
	EndIf

	If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOWriter_NumStyleExists($oDoc, $sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oObj.NumberingStyleName = $sNumStyle
	$iError = (__LOWriter_NumberingStyleCompare($oDoc, $oObj.NumberingStyleName(), $sNumStyle)) ? ($iError) : (BitOR($iError, 1))

	Return ($iError = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
EndFunc   ;==>_LOWriter_NumStyleCurrent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleCustomize
; Description ...: Retrieve and Set Numbering Style Customize settings. See Remarks.
; Syntax ........: _LOWriter_NumStyleCustomize(ByRef $oDoc, $oNumStyle, $iLevel[, $iNumFormat = Null[, $iStartAt = Null[, $sCharStyle = Null[, $iSubLevels = Null[, $sSepBefore = Null[, $sSepAfter = Null[, $bConsecutiveNum = Null[, $sBulletFont = Null[, $iCharDecimal = Null]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A Numbering Style object returned by a previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
;                  $iLevel              - an integer value (0-10). The Numbering Level to modify; enter 0 to modify all levels.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering scheme for the selected levels. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iStartAt            - [optional] an integer value. Default is Null. A new starting number for the current level
;                  $sCharStyle          - [optional] a string value. Default is Null. The Character Style that you want to use in an ordered list.
;                  $iSubLevels          - [optional] an integer value (1-10). Default is Null. Enter the number of previous levels to include in the outline format. See remarks.
;                  $sSepBefore          - [optional] a string value. Default is Null. A character or the text to display in front of the number in the list
;                  $sSepAfter           - [optional] a string value. Default is Null. A character or the text to display behind the number in the list.
;                  $bConsecutiveNum     - [optional] a boolean value. Default is Null. Increases the numbering by one as you go down each level in the list hierarchy.
;                  $sBulletFont         - [optional] a string value. Default is Null. The font to use for special characters that are associated with it. Note: $iNumFormat must be set to $LOW_NUM_STYLE_CHAR_SPECIAL(6) before these can be set.
;                  $iCharDecimal        - [optional] an integer value. Default is Null. The decimal value of the desired character. Note: $iNumFormat must be set to $LOW_NUM_STYLE_CHAR_SPECIAL(6) before these can be set.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iLevel not between 0 - 10.
;                  @Error 1 @Extended 5 Return 0 = $iNumFormat not an Integer, less than 0 or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iStartAt not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $sCharStyle not a string.
;                  @Error 1 @Extended 8 Return 0 = Character Style called in $sCharStyle, not found in document.
;                  @Error 1 @Extended 9 Return 0 = $iSubLevels not an Integer, less than 1 or greater than 10.
;                  @Error 1 @Extended 10 Return 0 = $iLevel called with 0 (all levels) And $iSubLevels greater than 1.
;                  @Error 1 @Extended 11 Return 0 = $iSubLevels greater than $iLevel.
;                  @Error 1 @Extended 12 Return 0 = $sSepBefore not a string.
;                  @Error 1 @Extended 13 Return 0 = $sSepAfter not a string.
;                  @Error 1 @Extended 14 Return 0 = $bConsecutiveNum not a Boolean.
;                  @Error 1 @Extended 15 Return 0 = $sBulletFont not a string.
;                  @Error 1 @Extended 16 Return 0 = Font style called in $sBulletFont not found in document.
;                  @Error 1 @Extended 17 Return 0 = $sBulletFont was called and Number Format not set to $LOW_NUM_STYLE_CHAR_SPECIAL.
;                  @Error 1 @Extended 18 Return 0 = $iCharDecimal not an Integer.
;                  @Error 1 @Extended 19 Return 0 = $iCharDecimal was called and Number Format not set to $LOW_NUM_STYLE_CHAR_SPECIAL.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error mapping setting values.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Numbering Rules Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Numbering Rule Array for level.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iNumFormat
;                  |                               2 = Error setting $iStartAt
;                  |                               4 = Error setting $sCharStyle
;                  |                               8 = Error setting $iSubLevels
;                  |                               16 = Error setting $sSepBefore
;                  |                               32 = Error setting $sSepAfter
;                  |                               64 = Error setting $bConsecutiveNum
;                  |                               128 = Error setting $sBulletFont
;                  |                               256 = Error setting $iCharDecimal
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully set the requested Properties.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 7 or 9 Element Array with values in order of function parameters. See remarks.
;                  @Error 0 @Extended 2 Return Array = Success. All optional parameters were called with Null, returning a 10 Element Array containing arrays of settings for each Numbering level corresponding to their position in the array. Each array will be as described above. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function should work just fine as the others do for modifying styles, but for setting Numbering Style settings, it would seem that the Array of Setting Objects passed by AutoIt is not recognized as an appropriate array/sequence by LibreOffice, and consequently causes a com.sun.star.lang.IllegalArgumentException COM error. See __LOWriter_NumStyleModify function for a more detailed explanation. This function can still be used to set and retrieve, setting values, however now, this function either inserts a temporary macro into $oDoc for performing the needed procedure, or if that fails, it invisibly opens an .odt Libre document and inserts a macro, see __LOWriter_NumStyleInitiateDocument which is then called with the necessary parameters to set.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  You can request setting values for one numbering level at a time, or all at once (see below). If Current numbering type is set to Bullet, the returned array will contain 9 elements, in the order of parameters, if the current numbering type is other than bullet style, a 7 element array will be returned, with the last two parameters excluded.
;                  If you retrieve the current settings for all levels (by calling $iLevel with 0), the return will be a 10 element array containing an array of settings for each Numbering Level.
;                  Call any optional parameter with Null keyword to skip it.
;                  When a lot of settings are set, especially for all levels, this function can be a bit slow.
;                  For $iSubLevels, if you enter "2" and the previous level uses the "A, B, C..." numbering scheme, the numbering scheme for the current level becomes: "A.1". The Maximum value, if $iLevel is set to 0, is 1.
; Related .......: _LOWriter_NumStyleCreate, _LOWriter_NumStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleCustomize(ByRef $oDoc, $oNumStyle, $iLevel, $iNumFormat = Null, $iStartAt = Null, $sCharStyle = Null, $iSubLevels = Null, $sSepBefore = Null, $sSepAfter = Null, $bConsecutiveNum = Null, $sBulletFont = Null, $iCharDecimal = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumRules
	Local $iError = 0
	Local $avCustomize[7], $aaAllLevels[10]
	Local $atNumLevel[0]
	Local $mNumLevel[]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oNumStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iLevel, 0, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iLevel = ($iLevel - 1) ; Numbering Levels are  0 based, minus 1 to compensate.

	$oNumRules = $oNumStyle.NumberingRules()
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iNumFormat, $iStartAt, $sCharStyle, $iSubLevels, $sSepBefore, $sSepAfter, $bConsecutiveNum, $sBulletFont, $iCharDecimal) Then
		For $i = (($iLevel = -1) ? (0) : ($iLevel)) To (($iLevel = -1) ? (9) : ($iLevel)) ; Determine if I'm retrieving settings for all levels or just one.
			$atNumLevel = $oNumRules.getByIndex($i)
			If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel)
			If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			If MapExists($mNumLevel, "BulletFont") Then
				__LO_ArrayFill($avCustomize, $atNumLevel[$mNumLevel["NumberingType"]].Value(), _
						$atNumLevel[$mNumLevel["StartWith"]].Value(), _
						$atNumLevel[$mNumLevel["CharStyleName"]].Value(), _
						$atNumLevel[$mNumLevel["ParentNumbering"]].Value(), _
						$atNumLevel[$mNumLevel["Prefix"]].Value(), _
						$atNumLevel[$mNumLevel["Suffix"]].Value(), _
						$oNumRules.IsContinuousNumbering(), _
						$atNumLevel[$mNumLevel["BulletFont"]].Value(), _
						Asc($atNumLevel[$mNumLevel["BulletChar"]].Value()))

			Else ; If not set for Bullet style, return only these settings as the rest don't exist.
				__LO_ArrayFill($avCustomize, $atNumLevel[$mNumLevel["NumberingType"]].Value(), _
						$atNumLevel[$mNumLevel["StartWith"]].Value(), _
						$atNumLevel[$mNumLevel["CharStyleName"]].Value(), _
						$atNumLevel[$mNumLevel["ParentNumbering"]].Value(), _
						$atNumLevel[$mNumLevel["Prefix"]].Value(), _
						$atNumLevel[$mNumLevel["Suffix"]].Value(), _
						$oNumRules.IsContinuousNumbering())
			EndIf

			If ($iLevel = -1) Then $aaAllLevels[$i] = $avCustomize
		Next

		Return ($iLevel = -1) ? (SetError($__LO_STATUS_SUCCESS, 2, $aaAllLevels)) : (SetError($__LO_STATUS_SUCCESS, 1, $avCustomize))
	EndIf

	For $i = (($iLevel = -1) ? (0) : ($iLevel)) To (($iLevel = -1) ? (9) : ($iLevel)) ; Determine if I am setting settings for all levels or just one.
		$atNumLevel = $oNumRules.getByIndex($i)
		If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel) ; Map what elements each setting is located at.
		If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		If ($iNumFormat <> Null) Then
			If Not __LO_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$atNumLevel[$mNumLevel["NumberingType"]].Value = $iNumFormat

			__LOWriter_NumStyleModify($oDoc, $oNumRules, $i, $atNumLevel) ; Modify the Setting in case it is switching from/to a bullet type.

			$atNumLevel = $oNumRules.getByIndex($i)
			If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel)
			If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		EndIf

		If ($iStartAt <> Null) Then
			If Not IsInt($iStartAt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$atNumLevel[$mNumLevel["StartWith"]].Value = $iStartAt
		EndIf

		If ($sCharStyle <> Null) Then
			If Not IsString($sCharStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
			If Not _LOWriter_CharStyleExists($oDoc, $sCharStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$atNumLevel[$mNumLevel["CharStyleName"]].Value = $sCharStyle
		EndIf

		If ($iSubLevels <> Null) Then
			If Not __LO_IntIsBetween($iSubLevels, 1, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			If ($iLevel = -1) And ($iSubLevels > 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0) ; -1 for $iLevel = 0 = Modify all Numbering Style levels.

			If ($iLevel <> -1) And ($iLevel < $iSubLevels) Then SetError($__LO_STATUS_INPUT_ERROR, 11, 0) ; Sub-level higher than requested level
			$atNumLevel[$mNumLevel["ParentNumbering"]].Value = $iSubLevels

			; If Document has "ListFormat" setting (Libre 7.2 +), Sub Levels ("ParentNumbering") wont accept a setting without also setting "List format", which means combining the corresponding "ListFormat"  number values + Prefix & Suffix.
			If MapExists($mNumLevel, "ListFormat") Then ; Test if "ListFormat" exists in the Numbering Rules.
				$atNumLevel[$mNumLevel["ListFormat"]].Value = __LOWriter_NumStyleListFormat($atNumLevel, $i, $iSubLevels)
			EndIf
		EndIf

		If ($sSepBefore <> Null) Then
			If Not IsString($sSepBefore) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

			$atNumLevel[$mNumLevel["Prefix"]].Value = $sSepBefore
			If MapExists($mNumLevel, "ListFormat") Then
				$atNumLevel[$mNumLevel["ListFormat"]].Value = __LOWriter_NumStyleListFormat($atNumLevel, $i, $atNumLevel[$mNumLevel["ParentNumbering"]].Value(), $sSepBefore, Null) ; Add prefix to ListFormat
			EndIf
		EndIf

		If ($sSepAfter <> Null) Then
			If Not IsString($sSepAfter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

			$atNumLevel[$mNumLevel["Suffix"]].Value = $sSepAfter
			If MapExists($mNumLevel, "ListFormat") Then
				$atNumLevel[$mNumLevel["ListFormat"]].Value = __LOWriter_NumStyleListFormat($atNumLevel, $i, $atNumLevel[$mNumLevel["ParentNumbering"]].Value(), Null, $sSepAfter) ; Add suffix to ListFormat
			EndIf
		EndIf

		If ($bConsecutiveNum <> Null) Then
			If Not IsBool($bConsecutiveNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)

			$oNumRules.IsContinuousNumbering = $bConsecutiveNum
		EndIf

		If ($sBulletFont <> Null) Then
			If Not IsString($sBulletFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)
			If Not _LOWriter_FontExists($sBulletFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)
			If Not MapExists($mNumLevel, "BulletFontName") Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)

			$atNumLevel[$mNumLevel["BulletFontName"]].Value = $sBulletFont
		EndIf

		If ($iCharDecimal <> Null) Then
			If Not IsInt($iCharDecimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 18, 0)
			If Not MapExists($mNumLevel, "BulletChar") Then Return SetError($__LO_STATUS_INPUT_ERROR, 19, 0)

			$atNumLevel[$mNumLevel["BulletChar"]].Value = Chr($iCharDecimal)
		EndIf

		__LOWriter_NumStyleModify($oDoc, $oNumRules, $i, $atNumLevel)
		$oNumStyle.NumberingRules = $oNumRules

		$atNumLevel = $oNumStyle.NumberingRules.getByIndex($i)
		If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel)
		If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		; Error Checking
		$iError = (__LO_VarsAreNull($iNumFormat)) ? ($iError) : (($atNumLevel[$mNumLevel["NumberingType"]].Value() = $iNumFormat) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iStartAt)) ? ($iError) : (($atNumLevel[$mNumLevel["StartWith"]].Value() = $iStartAt) ? ($iError) : (BitOR($iError, 2)))
		$iError = (__LO_VarsAreNull($sCharStyle)) ? ($iError) : ((__LOWriter_CharacterStyleCompare($oDoc, $atNumLevel[$mNumLevel["CharStyleName"]].Value(), $sCharStyle)) ? ($iError) : (BitOR($iError, 4)))
		$iError = (__LO_VarsAreNull($iSubLevels)) ? ($iError) : (($atNumLevel[$mNumLevel["ParentNumbering"]].Value() = $iSubLevels) ? ($iError) : (BitOR($iError, 8)))
		$iError = (__LO_VarsAreNull($sSepBefore)) ? ($iError) : (($atNumLevel[$mNumLevel["Prefix"]].Value() = $sSepBefore) ? ($iError) : (BitOR($iError, 16)))
		$iError = (__LO_VarsAreNull($sSepAfter)) ? ($iError) : (($atNumLevel[$mNumLevel["Suffix"]].Value() = $sSepAfter) ? ($iError) : (BitOR($iError, 32)))
		$iError = (__LO_VarsAreNull($bConsecutiveNum)) ? ($iError) : (($oNumStyle.NumberingRules.IsContinuousNumbering = $bConsecutiveNum) ? ($iError) : (BitOR($iError, 64)))
		$iError = (__LO_VarsAreNull($sBulletFont)) ? ($iError) : (($atNumLevel[$mNumLevel["BulletFont"]].Value.Name() = $sBulletFont) ? ($iError) : (BitOR($iError, 128)))
		$iError = (__LO_VarsAreNull($iCharDecimal)) ? ($iError) : ((Asc($atNumLevel[$mNumLevel["BulletChar"]].Value()) = $iCharDecimal) ? ($iError) : (BitOR($iError, 256)))
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStyleCustomize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleDelete
; Description ...: Delete a User-Created Numbering Style from a Document.
; Syntax ........: _LOWriter_NumStyleDelete(ByRef $oDoc, $oNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A Numbering Style object returned by a previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving "Numbering Styles" Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Numbering Style Name.
;                  @Error 3 @Extended 3 Return 0 = $sNumStyle is not a User-Created Numbering Style and cannot be deleted.
;                  @Error 3 @Extended 4 Return 0 = $sNumStyle is in use and cannot be deleted.
;                  @Error 3 @Extended 5 Return 0 = $sNumStyle still exists after deletion attempt.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. $sNumStyle was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_NumStyleCreate, _LOWriter_NumStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleDelete(ByRef $oDoc, $oNumStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumStyles
	Local $sNumStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oNumStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oNumStyles = $oDoc.StyleFamilies().getByName("NumberingStyles")
	If Not IsObj($oNumStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sNumStyle = $oNumStyle.Name()
	If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If Not $oNumStyle.isUserDefined() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	If $oNumStyle.isInUse() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; If Style is in use return an error unless force delete is true.

	$oNumStyles.removeByName($sNumStyle)

	Return ($oNumStyles.hasByName($sNumStyle)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleExists
; Description ...: Check whether a specified Numbering Style is available in a Document to use.
; Syntax ........: _LOWriter_NumStyleExists(ByRef $oDoc, $sNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNumStyle           - a string value. a Numbering Style name to search for.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sNumStyle not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if Numbering Style exists in the document, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleExists(ByRef $oDoc, $sNumStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oDoc.StyleFamilies.getByName("NumberingStyles").hasByName($sNumStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_NumStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleGetObj
; Description ...: Retrieve a Numbering Style Style Object for use with other Numbering Style Style functions.
; Syntax ........: _LOWriter_NumStyleGetObj(ByRef $oDoc, $sNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNumStyle           - a string value. The Numbering Style Style name to retrieve the Object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sNumStyle not a String.
;                  @Error 1 @Extended 3 Return 0 = Numbering Style Style called in $sNumStyle not found in Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Numbering Style Style Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Numbering Style Style successfully retrieved, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_NumStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleGetObj(ByRef $oDoc, $sNumStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_NumStyleExists($oDoc, $sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oNumStyle = $oDoc.StyleFamilies().getByName("NumberingStyles").getByName($sNumStyle)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNumStyle)
EndFunc   ;==>_LOWriter_NumStyleGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Numbering Style.
; Syntax ........: _LOWriter_NumStyleOrganizer(ByRef $oDoc, $oNumStyle[, $sNewNumStyleName = Null[, $bHidden = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A Numbering Style object returned by a previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
;                  $sNewNumStyleName    - [optional] a string value. Default is Null. The new name to set the Numbering Style called in $oNumStyle to.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, hide the style in the UI.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;                  @Error 1 @Extended 4 Return 0 = $sNewNumStyleName not a String.
;                  @Error 1 @Extended 5 Return 0 = Numbering Style name called in $sNewNumStyleName already exists in document.
;                  @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sNewParStyleName
;                  |                               2 = Error setting $bHidden
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 1 or 2 Element Array with values in order of function parameters. If the Libre Office version is below 4.0, the Array will contain 1 element because $bHidden is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_NumStyleCreate, _LOWriter_NumStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleOrganizer(ByRef $oDoc, $oNumStyle, $sNewNumStyleName = Null, $bHidden = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOrganizer[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oNumStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($sNewNumStyleName, $bHidden) Then
		If __LO_VersionCheck(4.0) Then
			__LO_ArrayFill($avOrganizer, $oNumStyle.Name(), $oNumStyle.Hidden())

		Else
			__LO_ArrayFill($avOrganizer, $oNumStyle.Name())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewNumStyleName <> Null) Then
		If Not IsString($sNewNumStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If _LOWriter_NumStyleExists($oDoc, $sNewNumStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oNumStyle.Name = $sNewNumStyleName
		$iError = (__LOWriter_NumberingStyleCompare($oDoc, $oNumStyle.Name(), $sNewNumStyleName)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LO_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oNumStyle.Hidden = $bHidden
		$iError = ($oNumStyle.Hidden() = $bHidden) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStylePosition
; Description ...: Retrieve and Set Numbering Style Position settings. See Remarks.
; Syntax ........: _LOWriter_NumStylePosition(ByRef $oDoc, $oNumStyle, $iLevel[, $iAlignedAt = Null[, $iNumAlign = Null[, $iFollowedBy = Null[, $iTabstop = Null[, $iIndent = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A Numbering Style object returned by a previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
;                  $iLevel              - an integer value (0-10). The Numbering Level to modify; enter 0 to modify all of them.
;                  $iAlignedAt          - [optional] an integer value. Default is Null. Specifies the first line indent. Set in Hundredths of a Millimeter (HMM).
;                  $iNumAlign           - [optional] an integer value (1-3). Default is Null. The alignment of the numbering symbols, in comparison to the "Aligned at" position. See Constants. $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iFollowedBy         - [optional] an integer value (0-2). Default is Null. Select the element that will follow the numbering: a tab stop, a space, or nothing; See Constants, $LOW_FOLLOW_BY_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iTabstop            - [optional] an integer value. Default is Null. If you select a tab stop to follow the numbering, you can enter a positive value as the tab stop position. Set in Hundredths of a Millimeter (HMM).
;                  $iIndent             - [optional] an integer value. Default is Null. Enter the distance from the left page margin to the start of all lines in the numbered paragraph that follow the first line. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;                  @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;                  @Error 1 @Extended 4 Return 0 = $iLevel not between 0 - 10.
;                  @Error 1 @Extended 5 Return 0 = $iAlignedAt not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iNumAlign not an Integer, less than 1 or greater than 3. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $iFollowedBy not an Integer, less than 0 or greater than 2. See Constants, $LOW_FOLLOW_BY_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iTabstop not an Integer.
;                  @Error 1 @Extended 9 Return 0 = $iIndent not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error mapping setting values.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Numbering Rules Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Numbering Rule Array for level.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAlignedAt
;                  |                               2 = Error setting $iNumAlign
;                  |                               4 = Error setting $iFollowedBy
;                  |                               8 = Error setting $iTabStop
;                  |                               16 = Error setting $iIndent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully set the requested Properties.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were called with Null, returning current settings in a 5 Element Array with values in order of function parameters.
;                  @Error 0 @Extended 2 Return Array = Success. All optional parameters were called with Null, returning a 10 Element Array containing arrays of settings for each Numbering level corresponding to their position in the array. Each array will be as described above. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function should work just fine as the others do for modifying styles, but for setting Numbering Style settings, it would seem that the Array of Setting Objects passed by AutoIt is not recognized as an appropriate array/Sequence by LibreOffice, and consequently causes a com.sun.star.lang.IllegalArgumentException COM error. See __LOWriter_NumStyleModify function for a more detailed explanation. This function can still be used to set and retrieve, setting values, however now, this function either inserts a temporary macro into $oDoc for performing the needed procedure, or if that fails, it invisibly opens an .odt Libre document and inserts a macro, (see __LOWriter_NumStyleInitiateDocument), which is then called with the necessary parameters to set.
;                  Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings. You can request setting values for one numbering level at a time, or all at once (see below).
;                  If you retrieve the current settings for all levels (by calling $iLevel with 0), the return will be a 10 element array containing an array of settings for each Numbering Level.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_NumStyleCreate, _LOWriter_NumStyleGetObj, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStylePosition(ByRef $oDoc, $oNumStyle, $iLevel, $iAlignedAt = Null, $iNumAlign = Null, $iFollowedBy = Null, $iTabStop = Null, $iIndent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumRules
	Local $iError = 0
	Local $avPosition[5], $aaAllLevels[10]
	Local $atNumLevel[0]
	Local $mNumLevel[]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oNumStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iLevel, 0, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iLevel = ($iLevel - 1) ; Numbering Levels are  0 based, minus 1 to compensate.

	$oNumRules = $oNumStyle.NumberingRules()
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iAlignedAt, $iNumAlign, $iFollowedBy, $iTabStop, $iIndent) Then
		For $i = (($iLevel = -1) ? (0) : ($iLevel)) To (($iLevel = -1) ? (9) : ($iLevel)) ; Determine if I'm retrieving settings for all levels or just one.
			$atNumLevel = $oNumRules.getByIndex($i)
			If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel)
			If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			__LO_ArrayFill($avPosition, $atNumLevel[$mNumLevel["FirstLineIndent"]].Value(), _
					$atNumLevel[$mNumLevel["Adjust"]].Value(), _
					$atNumLevel[$mNumLevel["LabelFollowedBy"]].Value(), _
					$atNumLevel[$mNumLevel["ListtabStopPosition"]].Value(), _
					$atNumLevel[$mNumLevel["IndentAt"]].Value())

			If ($iLevel = -1) Then $aaAllLevels[$i] = $avPosition
		Next

		Return ($iLevel = -1) ? (SetError($__LO_STATUS_SUCCESS, 2, $aaAllLevels)) : (SetError($__LO_STATUS_SUCCESS, 1, $avPosition))
	EndIf

	For $i = (($iLevel = -1) ? (0) : ($iLevel)) To (($iLevel = -1) ? (9) : ($iLevel)) ; Determine if I am setting settings for all levels or just one.
		$atNumLevel = $oNumRules.getByIndex($i)
		If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel) ; Map what elements each setting is located at.
		If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		If ($iAlignedAt <> Null) Then
			If Not IsInt($iAlignedAt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$atNumLevel[$mNumLevel["FirstLineIndent"]].Value = $iAlignedAt
		EndIf

		If ($iNumAlign <> Null) Then
			If Not __LO_IntIsBetween($iNumAlign, $LOW_ORIENT_HORI_RIGHT, $LOW_ORIENT_HORI_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$atNumLevel[$mNumLevel["Adjust"]].Value = $iNumAlign
		EndIf

		If ($iFollowedBy <> Null) Then
			If Not __LO_IntIsBetween($iFollowedBy, $LOW_FOLLOW_BY_TABSTOP, $LOW_FOLLOW_BY_NEWLINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$atNumLevel[$mNumLevel["LabelFollowedBy"]].Value = $iFollowedBy
		EndIf

		If ($iTabStop <> Null) Then
			If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$atNumLevel[$mNumLevel["ListtabStopPosition"]].Value = $iTabStop
		EndIf

		If ($iIndent <> Null) Then
			If Not IsInt($iIndent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

			$atNumLevel[$mNumLevel["IndentAt"]].Value = $iIndent
		EndIf

		__LOWriter_NumStyleModify($oDoc, $oNumRules, $i, $atNumLevel)
		$oNumStyle.NumberingRules = $oNumRules

		$atNumLevel = $oNumStyle.NumberingRules.getByIndex($i)
		If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel)
		If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		; Error Checking:

		$iError = (__LO_VarsAreNull($iAlignedAt)) ? ($iError) : ((__LO_IntIsBetween($atNumLevel[$mNumLevel["FirstLineIndent"]].Value(), $iAlignedAt - 1, $iAlignedAt + 1)) ? ($iError) : (BitOR($iError, 1)))
		$iError = (__LO_VarsAreNull($iNumAlign)) ? ($iError) : (($atNumLevel[$mNumLevel["Adjust"]].Value() = $iNumAlign) ? ($iError) : (BitOR($iError, 2)))
		$iError = (__LO_VarsAreNull($iFollowedBy)) ? ($iError) : (($atNumLevel[$mNumLevel["LabelFollowedBy"]].Value() = $iFollowedBy) ? ($iError) : (BitOR($iError, 4)))
		$iError = (__LO_VarsAreNull($iTabStop)) ? ($iError) : ((__LO_IntIsBetween($atNumLevel[$mNumLevel["ListtabStopPosition"]].Value(), $iTabStop - 1, $iTabStop + 1)) ? ($iError) : (BitOR($iError, 8)))
		$iError = (__LO_VarsAreNull($iIndent)) ? ($iError) : ((__LO_IntIsBetween($atNumLevel[$mNumLevel["IndentAt"]].Value(), $iIndent - 1, $iIndent + 1)) ? ($iError) : (BitOR($iError, 16)))
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStylePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleSetLevel
; Description ...: Set the numbering style level for a paragraph by Cursor or paragraph Object.
; Syntax ........: _LOWriter_NumStyleSetLevel(ByRef $oObj[, $iLevel = Null])
; Parameters ....: $oObj                - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object returned from _LOWriter_ParObjCreateList function.
;                  $iLevel              - [optional] an integer value (1-10). Default is Null. The Numbering Style level to set the paragraph to.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj does not support Paragraph Properties Service.
;                  @Error 1 @Extended 3 Return 0 = $iLevel not an Integer, less than 1 or greater than 10.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLevel
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Numbering Style successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParObjCreateList, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleSetLevel(ByRef $oObj, $iLevel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iLevel, 1, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iLevel) Then Return SetError($__LO_STATUS_SUCCESS, 1, ($oObj.NumberingLevel() + 1)) ; Plus one to compensate for Levels being 0 Based.

	$iLevel -= 1 ; Level is 0 Based, minus one to compensate.

	$oObj.NumberingLevel = $iLevel

	Return ($oObj.NumberingLevel() = $iLevel) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_NumStyleSetLevel

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStylesGetNames
; Description ...: Retrieve an array of all Numbering Style names available for a document.
; Syntax ........: _LOWriter_NumStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False[, $bDisplayName = False]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True only User-Created Numbering Styles are returned.
;                  $bAppliedOnly        - [optional] a boolean value. Default is False. If True only Applied Numbering Styles are returned.
;                  $bDisplayName        - [optional] a boolean value. Default is False. If True, the style name displayed in the UI (Display Name), instead of the programmatic style name, is returned. See remarks.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bUserOnly not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bAppliedOnly not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bDisplayName not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Numbering Style names.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. An Array containing all Numbering Styles matching the called parameters. See remarks. @Extended contains the count of results returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If Only a Document object is input, all available Numbering styles will be returned.
;                  If Both $bUserOnly and $bAppliedOnly are called with True, only User-Created styles that are applied are returned.
;                  Five Numbering styles have different internal names:
;                  - "Bullet •" is internally called List 1.
;                  - "Bullet –" is internally called List 2.
;                  - "Bullet [Checkbox]" is internally called List 3.
;                  - "Bullet [Arrow]" is internally called List 4.
;                  - "Bullet [X]" is internally called List 5.
;                  Previous to LibreOffice 25.2 either name would work when setting a Style, however after 25.2 only the internal, or programmatic style names, will work.
;                  Calling $bDisplayName with True will return a list of Style names, as the user sees them in the UI, in the same order as they are returned if $bDisplayName is False. It is best not to use these when setting Styling.
;                  Numbering Styles with special characters may return a name like "Bullet ?" when $bDisplayName is True.
; Related .......: _LOWriter_NumStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False, $bDisplayName = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asStyles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bDisplayName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$asStyles = __LO_StylesGetNames($oDoc, "NumberingStyles", $bUserOnly, $bAppliedOnly, $bDisplayName)
	If Not IsArray($asStyles) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, UBound($asStyles), $asStyles)
EndFunc   ;==>_LOWriter_NumStylesGetNames
