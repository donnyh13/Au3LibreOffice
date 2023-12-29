#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

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
; _LOWriter_NumStyleCustomize
; _LOWriter_NumStyleDelete
; _LOWriter_NumStyleExists
; _LOWriter_NumStyleGetObj
; _LOWriter_NumStyleOrganizer
; _LOWriter_NumStylePosition
; _LOWriter_NumStyleSet
; _LOWriter_NumStyleSetLevel
; _LOWriter_NumStylesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleCreate
; Description ...: Create a new Numbering Style in a Document.
; Syntax ........: _LOWriter_NumStyleCreate(ByRef $oDoc, $sNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNumStyle           - a string value. The Name of the new Numbering Style to create.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sNumStyle not a String.
;				   @Error 1 @Extended 3 Return 0 = Numbering Style name called in $sNumStyle already exists in this document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Retrieving "NumberingStyle" Object.
;				   @Error 2 @Extended 2 Return 0 = Error Creating "com.sun.star.style.NumberingStyle" Object.
;				   @Error 2 @Extended 3 Return 0 = Error Retrieving New Numbering Style Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error creating new Numbering Style by Name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. New Numbering Style successfully created. Returning Numbering Style Object.
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
	If Not IsObj($oNumStyles) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	If _LOWriter_NumStyleExists($oDoc, $sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$oStyle = $oDoc.createInstance("com.sun.star.style.NumberingStyle")
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oNumStyles.insertByName($sNumStyle, $oStyle)

	If Not $oNumStyles.hasByName($sNumStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oNumStyle = $oNumStyles.getByName($sNumStyle)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNumStyle)
EndFunc   ;==>_LOWriter_NumStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleCustomize
; Description ...: Retrieve and Set Numbering Style Customize settings. See Remarks.
; Syntax ........: _LOWriter_NumStyleCustomize(ByRef $oDoc, $oNumStyle, $iLevel[, $iNumFormat = Null[, $iStartAt = Null[, $sCharStyle = Null[, $iSubLevels = Null[, $sSepBefore = Null[, $sSepAfter = Null[, $bConsecutiveNum = Null[, $sBulletFont = Null[, $iCharDecimal = Null]]]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A NumberingStyle object returned by previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
;                  $iLevel              - an integer value (0-10). The Numbering Level to modify; enter 0 to modify all levels.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering scheme for the selected levels. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iStartAt            - [optional] an integer value. Default is Null. A new starting number for the current level
;                  $sCharStyle          - [optional] a string value. Default is Null. The Character Style that you want to use in an ordered list.
;                  $iSubLevels          - [optional] an integer value (1-10). Default is Null. Enter the number of previous levels to include in the outline format. For example, if you enter "2" and the previous level uses the "A, B, C..." numbering scheme, the numbering scheme for the current level becomes: "A.1". Maximum value, if $iLevel is set to 0, is 1.
;                  $sSepBefore          - [optional] a string value. Default is Null. A character or the text to display in front of the number in the list
;                  $sSepAfter           - [optional] a string value. Default is Null. A character or the text to display behind the number in the list.
;                  $bConsecutiveNum     - [optional] a boolean value. Default is Null. Increases the numbering by one as you go down each level in the list hierarchy.
;                  $sBulletFont         - [optional] a string value. Default is Null. The font to use for special characters that are associated with it. Note: $iNumFormat must be set to $LOW_NUM_STYLE_CHAR_SPECIAL(6) before these can be set.
;                  $iCharDecimal        - [optional] an integer value. Default is Null. The decimal value of the desired character. Note: $iNumFormat must be set to $LOW_NUM_STYLE_CHAR_SPECIAL(6) before these can be set.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;				   @Error 1 @Extended 4 Return 0 = $iLevel not between 0 - 10.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormat not an integer, less than 0, or greater than 71, see Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $iStartAt not an integer.
;				   @Error 1 @Extended 7 Return 0 = $sCharStyle not a string.
;				   @Error 1 @Extended 8 Return 0 = Character Style called in $sCharStyle, not found in document.
;				   @Error 1 @Extended 9 Return 0 = $iSubLevels not an Integer, less than 1, or greater than 10.
;				   @Error 1 @Extended 10 Return 0 = $iLevel set to 0 (all levels) And $iSubLevels greater than 1.
;				   @Error 1 @Extended 11 Return 0 = $iSubLevels greater than $iLevel.
;				   @Error 1 @Extended 12 Return 0 = $sSepBefore not a string.
;				   @Error 1 @Extended 13 Return 0 = $sSepAfter not a string.
;				   @Error 1 @Extended 14 Return 0 = $bConsecutiveNum not a Boolean.
;				   @Error 1 @Extended 15 Return 0 = $sBulletFont not a string.
;				   @Error 1 @Extended 16 Return 0 = Font style called in $sBulletFont not found in document.
;				   @Error 1 @Extended 17 Return 0 = $iCharDecimal not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving NumberingRules Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving NumberingRules Object for error checking.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving current settings, $iLevel set to 0, cannot retrieve settings for more than one level at a time.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iStartAt
;				   |								4 = Error setting $sCharStyle
;				   |								8 = Error setting $iSubLevels
;				   |								16 = Error setting $sSepBefore
;				   |								32 = Error setting $sSepAfter
;				   |								64 = Error setting $bConsecutiveNum
;				   |								128 = Error setting $sBulletFont
;				   |								256 = Error setting $iCharDecimal
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully set the requested Properties.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 7 or 9 Element Array with values in order of function parameters. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function should work just fine as the others do for modifying styles, but for setting Numbering Style
;						settings, it would seem that the Array of Setting Objects passed by AutoIt is not recognized as an
;						appropriate array/sequence by LibreOffice, and consequently causes a
;						com.sun.star.lang.IllegalArgumentException COM error. See __LOWriter_NumStyleModify function for a more
;						detailed explanation. This function can still be used to set and retrieve, setting values, however now,
;						this function either inserts a temporary macro into $oDoc for performing the needed procedure, or if
;						that fails, it invisibly opens an .odt Libre document and inserts a macro, see
;						__LOWriter_NumStyleInitiateDocument which is then called with the necessary parameters to set.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings. Note: Can only request setting values for one numbering level at a time, you
;						aren't able to call $iLevel with 0 to retrieve all at once. If Current numbering type is set to Bullet,
;						the returned array will contain 9 elements, in the order of parameters, if the current numbering type is
;						other than bullet style, a 7 element array will be returned, with the last two parameters excluded.
;				   Call any optional parameter with Null keyword to skip it.
;				   When a lot of settings are set, especially for all levels, this function can be a bit slow.
; Related .......: _LOWriter_NumStyleCreate, _LOWriter_NumStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleCustomize(ByRef $oDoc, $oNumStyle, $iLevel, $iNumFormat = Null, $iStartAt = Null, $sCharStyle = Null, $iSubLevels = Null, $sSepBefore = Null, $sSepAfter = Null, $bConsecutiveNum = Null, $sBulletFont = Null, $iCharDecimal = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumRules
	Local $iError = 0
	Local $aNumSettings[9][2]
	Local $iRowCount = 0
	Local $avCustomize[7]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oNumStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_IntIsBetween($iLevel, 0, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	$iLevel = ($iLevel - 1) ; Numbering Levels are  0 based, minus 1 to compensate.

	$oNumRules = $oNumStyle.NumberingRules()
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumFormat, $iStartAt, $sCharStyle, $iSubLevels, $sSepBefore, $sSepAfter, $bConsecutiveNum, $sBulletFont, $iCharDecimal) Then
		If ($iLevel = -1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; can only get settings for one level at a time.

		If (__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "NumberingType") = $LOW_NUM_STYLE_CHAR_SPECIAL) Then
			__LOWriter_ArrayFill($avCustomize, __LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "NumberingType"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "StartWith"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "CharStyleName"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "ParentNumbering"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Prefix"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Suffix"), _
					$oNumStyle.NumberingRules.IsContinuousNumbering(), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "BulletFont").Name(), _
					Asc(__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "BulletChar")))

		Else ; If not set for Bullet style, return only these settings.
			__LOWriter_ArrayFill($avCustomize, __LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "NumberingType"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "StartWith"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "CharStyleName"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "ParentNumbering"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Prefix"), _
					__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Suffix"), _
					$oNumStyle.NumberingRules.IsContinuousNumbering())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avCustomize)
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$aNumSettings[$iRowCount][0] = "NumberingType"
		$aNumSettings[$iRowCount][1] = $iNumFormat
		$iRowCount += 1
	EndIf

	If ($iStartAt <> Null) Then
		If Not IsInt($iStartAt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$aNumSettings[$iRowCount][0] = "StartWith"
		$aNumSettings[$iRowCount][1] = $iStartAt
		$iRowCount += 1
	EndIf

	If ($sCharStyle <> Null) Then
		If Not IsString($sCharStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If Not _LOWriter_CharStyleExists($oDoc, $sCharStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$aNumSettings[$iRowCount][0] = "CharStyleName"
		$aNumSettings[$iRowCount][1] = $sCharStyle
		$iRowCount += 1
	EndIf

	If ($iSubLevels <> Null) Then
		If Not __LOWriter_IntIsBetween($iSubLevels, 1, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If ($iLevel = -1) And ($iSubLevels > 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0) ; -1 for $iLevel = 0 = Modify all Numbering Style levels.
		If ($iLevel <> -1) And ($iLevel < $iSubLevels) Then SetError($__LO_STATUS_INPUT_ERROR, 11, 0) ; Sub-level higher than requested level
		$aNumSettings[$iRowCount][0] = "ParentNumbering"
		$aNumSettings[$iRowCount][1] = $iSubLevels
		$iRowCount += 1

		; If Document has "ListFormat" setting (Libre 7.2 +), Sub Levels ("ParentNumbering") wont accept a setting without
		; also setting "List format", which means combining the corresponding "ListFormat"  number values + Prefix & Suffix.
		__LOWriter_NumStyleRetrieve($oNumRules, 9, "ListFormat") ; Test if "ListFormat" exists in the Numbering Rules.
		If (@error = 0) Then ;  If List Format does exist, modify it.
			$aNumSettings[$iRowCount][0] = "ListFormat"
			$aNumSettings[$iRowCount][1] = __LOWriter_NumStyleListFormat($oNumRules, $iLevel, $iSubLevels)
			If (@error = 0) Then $iRowCount += 1 ;If No errors count up, else just let it be overwritten.
		EndIf
	EndIf

	If ($sSepBefore <> Null) Then
		If Not IsString($sSepBefore) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)
		$aNumSettings[$iRowCount][0] = "Prefix"
		$aNumSettings[$iRowCount][1] = $sSepBefore
		$iRowCount += 1
	EndIf

	If ($sSepAfter <> Null) Then
		If Not IsString($sSepAfter) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)
		$aNumSettings[$iRowCount][0] = "Suffix"
		$aNumSettings[$iRowCount][1] = $sSepAfter
		$iRowCount += 1
	EndIf

	If ($bConsecutiveNum <> Null) Then
		If Not IsBool($bConsecutiveNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 14, 0)
		$oNumRules.IsContinuousNumbering = $bConsecutiveNum
	EndIf

	If ($sBulletFont <> Null) Then
		If Not IsString($sBulletFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 15, 0)
		If Not _LOWriter_FontExists($oDoc, $sBulletFont) Then Return SetError($__LO_STATUS_INPUT_ERROR, 16, 0)
		$aNumSettings[$iRowCount][0] = "BulletFontName"
		$aNumSettings[$iRowCount][1] = $sBulletFont
		$iRowCount += 1
	EndIf

	If ($iCharDecimal <> Null) Then
		If Not IsInt($iCharDecimal) Then Return SetError($__LO_STATUS_INPUT_ERROR, 17, 0)
		$aNumSettings[$iRowCount][0] = "BulletChar"
		$aNumSettings[$iRowCount][1] = Chr($iCharDecimal)
		$iRowCount += 1
	EndIf

	ReDim $aNumSettings[$iRowCount][2]

	__LOWriter_NumStyleModify($oDoc, $oNumRules, $iLevel, $aNumSettings)

	$oNumStyle.NumberingRules = $oNumRules

	$oNumRules = $oNumStyle.NumberingRules() ; Retrieve Numbering Rules a second time for error checking.
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	$iLevel = ($iLevel = -1) ? (9) : ($iLevel) ;If Level is set to -1 (modify all), set to last level to check the settings.

	; Error Checking
	$iError = ($iNumFormat = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "NumberingType") = $iNumFormat) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iStartAt = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "StartWith") = $iStartAt) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($sCharStyle = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "CharStyleName") = $sCharStyle) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iSubLevels = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "ParentNumbering") = $iSubLevels) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($sSepBefore = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Prefix") = $sSepBefore) ? ($iError) : (BitOR($iError, 16)))
	$iError = ($sSepAfter = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Suffix") = $sSepAfter) ? ($iError) : (BitOR($iError, 32)))
	$iError = ($bConsecutiveNum = Null) ? ($iError) : (($oNumStyle.NumberingRules.IsContinuousNumbering = $bConsecutiveNum) ? ($iError) : (BitOR($iError, 64)))
	$iError = ($sBulletFont = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "BulletFont").Name() = $sBulletFont) ? ($iError) : (BitOR($iError, 128)))
	$iError = ($iCharDecimal = Null) ? ($iError) : ((Asc(__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "BulletChar")) = $iCharDecimal) ? ($iError) : (BitOR($iError, 256)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStyleCustomize

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleDelete
; Description ...: Delete a User-Created Numbering Style from a Document.
; Syntax ........: _LOWriter_NumStyleDelete(ByRef $oDoc, $oNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A Numbering Style object returned by previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving "Numbering Styles" Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Numbering Style Name.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $sNumStyle is not a User-Created Numbering Style and cannot be deleted.
;				   @Error 3 @Extended 2 Return 0 = $sNumStyle is in use and cannot be deleted.
;				   @Error 3 @Extended 3 Return 0 = $sNumStyle still exists after deletion attempt.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. $sNumStyle was successfully deleted.
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
	If Not IsObj($oNumStyles) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$sNumStyle = $oNumStyle.Name()
	If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not $oNumStyle.isUserDefined() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oNumStyle.isInUse() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; If Style is in use return an error unless force delete is true.

	$oNumStyles.removeByName($sNumStyle)

	Return ($oNumStyles.hasByName($sNumStyle)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleExists
; Description ...: Check whether a specified Numbering Style is available in a Document to use.
; Syntax ........: _LOWriter_NumStyleExists(ByRef $oDoc, $sNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNumStyle           - a string value. a Numbering Style name to search for.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sNumStyle not a String.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean  = Success. Returns True if Numbering Style exists in the document, else False.
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
; Name ..........: _LOWriter_NumStyleStyleGetObj
; Description ...: Retrieve a Numbering Style Style Object for use with other  Numbering Style Style functions.
; Syntax ........: _LOWriter_NumStyleStyleGetObj(ByRef $oDoc, $sNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sNumStyle           - a string value. The Numbering Style Style name to retrieve the Object for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sNumStyle not a String.
;				   @Error 1 @Extended 3 Return 0 = Numbering Style Style called in $sNumStyle not found in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Numbering Style Style Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Numbering Style Style successfully retrieved, returning its Object.
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
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oNumStyle)
EndFunc   ;==>_LOWriter_NumStyleGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Numbering Style.
; Syntax ........: _LOWriter_NumStyleOrganizer(ByRef $oDoc, $oNumStyle[, $sNewNumStyleName = Null[, $bHidden = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A Numbering Style object returned by previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
;                  $sNewNumStyleName    - [optional] a string value. Default is Null. The new name to set the Numbering Style called in $oNumStyle to.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, hide the style in the UI.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;				   @Error 1 @Extended 4 Return 0 = $sNewNumStyleName not a String.
;				   @Error 1 @Extended 5 Return 0 = Numbering Style name called in $sNewNumStyleName already exists in document.
;				   @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sNewParStyleName
;				   |								2 = Error setting $bHidden
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 1 or 2 Element Array with values in order of function parameters. If the Libre Office version is below 4.0, the Array will contain 1 element because $bHidden is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
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

	If __LOWriter_VarsAreNull($sNewNumStyleName, $bHidden) Then
		If __LOWriter_VersionCheck(4.0) Then
			__LOWriter_ArrayFill($avOrganizer, $oNumStyle.Name(), $oNumStyle.Hidden())
		Else
			__LOWriter_ArrayFill($avOrganizer, $oNumStyle.Name())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewNumStyleName <> Null) Then
		If Not IsString($sNewNumStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If _LOWriter_NumStyleExists($oDoc, $sNewNumStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oNumStyle.Name = $sNewNumStyleName
		$iError = ($oNumStyle.Name() = $sNewNumStyleName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$oNumStyle.Hidden = $bHidden
		$iError = ($oNumStyle.Hidden() = $bHidden) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStylePosition
; Description ...: Retrieve and Set Numbering Style Position settings. See Remarks.
; Syntax ........: _LOWriter_NumStylePosition(ByRef $oDoc, $sNumStyle, $iLevel[, $iAlignedAt = Null[, $iNumAlign = Null[, $iFollowedBy = Null[, $iTabstop = Null[, $iIndent = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oNumStyle           - [in/out] an object. A Numbering Style object returned by previous _LOWriter_NumStyleCreate, or _LOWriter_NumStyleGetObj function.
;                  $iLevel              - an integer value (0-10). The Numbering Level to modify; enter 0 to modify all of them.
;                  $iAlignedAt          - [optional] an integer value. Default is Null. Specifies the first line indent. Set in Micrometers.
;                  $iNumAlign           - [optional] an integer value (1-3). Default is Null. The alignment of the numbering symbols, in comparison to the "Aligned at" position. See Constants. $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iFollowedBy         - [optional] an integer value (0-2). Default is Null. Select the element that will follow the numbering: a tab stop, a space, or nothing; See Constants, $LOW_FOLLOW_BY_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iTabstop            - [optional] an integer value. Default is Null. If you select a tab stop to follow the numbering, you can enter a positive value as the tab stop position. Set in Micrometers.
;                  $iIndent             - [optional] an integer value. Default is Null. Enter the distance from the left page margin to the start of all lines in the numbered paragraph that follow the first line. Set in Micrometers.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oNumStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oNumStyle not a Numbering Style Object.
;				   @Error 1 @Extended 4 Return 0 = $iLevel not between 0 - 10.
;				   @Error 1 @Extended 5 Return 0 = $iAlignedAt not an integer.
;				   @Error 1 @Extended 6 Return 0 = $iNumAlign not an integer, less than 1, or higher than 3. See Constants, $LOW_ORIENT_HORI_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $iFollowedBy not an integer, less than 0, or higher than 2. See Constants, $LOW_FOLLOW_BY_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $iTabstop not an Integer.
;				   @Error 1 @Extended 9 Return 0 = $iIndent not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving NumberingRules Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving NumberingRules Object for error checking.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving current settings, $iLevel set to 0, cannot retrieve settings for more than one level at a time.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iAlignedAt
;				   |								2 = Error setting $iNumAlign
;				   |								4 = Error setting $iFollowedBy
;				   |								8 = Error setting $iTabStop
;				   |								16 = Error setting $iIndent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully set the requested Properties.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function should work just fine as the others do for modifying styles, but for setting Numbering Style
;						settings, it would seem that the Array of Setting Objects passed by AutoIt is not recognized as an
;						appropriate array/Sequence by LibreOffice, and consequently causes a
;						com.sun.star.lang.IllegalArgumentException COM error. See __LOWriter_NumStyleModify function for a more
;						detailed explanation. This function can still be used to set and retrieve, setting values, however now,
;						this function either inserts a temporary macro into $oDoc for performing the needed procedure, or if
;						that fails, it invisibly opens an .odt Libre document and inserts a macro, (see
;						__LOWriter_NumStyleInitiateDocument), which is then called with the necessary parameters to set.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings. Note: You can only request setting values for one numbering level at a time, you aren't able to call $iLevel with 0 to retrieve all at once.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_NumStyleCreate, _LOWriter_NumStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStylePosition(ByRef $oDoc, $oNumStyle, $iLevel, $iAlignedAt = Null, $iNumAlign = Null, $iFollowedBy = Null, $iTabStop = Null, $iIndent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumRules
	Local $iError = 0
	Local $aNumSettings[5][2]
	Local $iRowCount = 0
	Local $avPosition[5]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oNumStyle.supportsService("com.sun.star.style.Style") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_IntIsBetween($iLevel, 0, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	$iLevel = ($iLevel - 1) ; Numbering Levels are  0 based, minus 1 to compensate.

	$oNumRules = $oNumStyle.NumberingRules()
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iAlignedAt, $iNumAlign, $iFollowedBy, $iTabStop, $iIndent) Then
		If ($iLevel = -1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; only get settings for one level at a time.
		__LOWriter_ArrayFill($avPosition, __LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "FirstLineIndent"), _
				__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Adjust"), _
				__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "LabelFollowedBy"), _
				__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "ListtabStopPosition"), _
				__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "IndentAt"))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf

	If ($iAlignedAt <> Null) Then
		If Not IsInt($iAlignedAt) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$aNumSettings[$iRowCount][0] = "FirstLineIndent"
		$aNumSettings[$iRowCount][1] = $iAlignedAt
		$iRowCount += 1
	EndIf

	If ($iNumAlign <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumAlign, $LOW_ORIENT_HORI_RIGHT, $LOW_ORIENT_HORI_LEFT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$aNumSettings[$iRowCount][0] = "Adjust"
		$aNumSettings[$iRowCount][1] = $iNumAlign
		$iRowCount += 1
	EndIf

	If ($iFollowedBy <> Null) Then
		If Not __LOWriter_IntIsBetween($iFollowedBy, $LOW_FOLLOW_BY_TABSTOP, $LOW_FOLLOW_BY_NEWLINE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$aNumSettings[$iRowCount][0] = "LabelFollowedBy"
		$aNumSettings[$iRowCount][1] = $iFollowedBy
		$iRowCount += 1
	EndIf

	If ($iTabStop <> Null) Then
		If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$aNumSettings[$iRowCount][0] = "ListtabStopPosition"
		$aNumSettings[$iRowCount][1] = $iTabStop
		$iRowCount += 1
	EndIf

	If ($iIndent <> Null) Then
		If Not IsInt($iIndent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		$aNumSettings[$iRowCount][0] = "IndentAt"
		$aNumSettings[$iRowCount][1] = $iIndent
		$iRowCount += 1
	EndIf

	ReDim $aNumSettings[$iRowCount][2]

	__LOWriter_NumStyleModify($oDoc, $oNumRules, $iLevel, $aNumSettings)

	$oNumStyle.NumberingRules = $oNumRules

	; Error Checking:
	$oNumRules = $oNumStyle.NumberingRules()
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
	$iLevel = ($iLevel = -1) ? (9) : ($iLevel) ;If Level is set to -1 (modify all), set to last level to check the settings.

	$iError = ($iAlignedAt = Null) ? ($iError) : ((__LOWriter_IntIsBetween(__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "FirstLineIndent"), $iAlignedAt - 1, $iAlignedAt + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iNumAlign = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Adjust") = $iNumAlign) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iFollowedBy = Null) ? ($iError) : ((__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "LabelFollowedBy") = $iFollowedBy) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iTabStop = Null) ? ($iError) : ((__LOWriter_IntIsBetween(__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "ListtabStopPosition"), $iTabStop - 1, $iTabStop + 1)) ? ($iError) : (BitOR($iError, 8)))
	$iError = ($iIndent = Null) ? ($iError) : ((__LOWriter_IntIsBetween(__LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "IndentAt"), $iIndent - 1, $iIndent + 1)) ? ($iError) : (BitOR($iError, 16)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_NumStylePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleSet
; Description ...: Set a numbering style for a paragraph by Cursor or paragraph Object.
; Syntax ........: _LOWriter_NumStyleSet(ByRef $oDoc, ByRef $oObj, $sNumStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object returned from _LOWriter_ParObjCreateList function.
;                  $sNumStyle           - a string value. The Numbering Style name to set the paragraph to.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oObj does not support Paragraph Properties Service.
;				   @Error 1 @Extended 4 Return 0 = $sNumStyle not a String.
;				   @Error 1 @Extended 5 Return 0 = Numbering Style called in $sNumStyle doesn't exist in Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Error setting Numbering Style.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Numbering Style successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParObjCreateList, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_NumStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleSet(ByRef $oDoc, ByRef $oObj, $sNumStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOWriter_NumStyleExists($oDoc, $sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	$oObj.NumberingStyleName = $sNumStyle
	Return ($oObj.NumberingStyleName() = $sNumStyle) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))

EndFunc   ;==>_LOWriter_NumStyleSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStyleSetLevel
; Description ...: Set the numbering style level for a paragraph by Cursor or paragraph Object.
; Syntax ........: _LOWriter_NumStyleSetLevel(ByRef $oDoc, ByRef $oObj[, $iLevel = Null])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object returned from _LOWriter_ParObjCreateList function.
;                  $iLevel              - [optional] an integer value (1-10). Default is Null. The Numbering Style level to set the paragraph to. Set to Null to retrieve the current level set.
; Return values .: Success: 1 or Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oObj does not support Paragraph Properties Service.
;				   @Error 1 @Extended 4 Return 0 = $iLevel not an Integer, less than 1, or greater than 10.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Error setting Numbering Style level.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Numbering Style successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParObjCreateList, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStyleSetLevel(ByRef $oDoc, ByRef $oObj, $iLevel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_IntIsBetween($iLevel, 1, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($iLevel = Null) Then Return SetError($__LO_STATUS_SUCCESS, 1, ($oObj.NumberingLevel() + 1)) ; Plus one to compensate for Levels being 0 Based.

	$iLevel -= 1 ;Level is 0 Based, minus one to compensate.

	$oObj.NumberingLevel = $iLevel

	Return ($oObj.NumberingLevel() = $iLevel) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_NumStyleSetLevel

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_NumStylesGetNames
; Description ...: Retrieve a list of all Numbering Style names available for a document.
; Syntax ........: _LOWriter_NumStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bUserOnly           - [optional] a boolean value. Default is False. If True only User-Created Numbering Styles are returned.
;                  $bAppliedOnly        - [optional] a boolean value. Default is False. If True only Applied Numbering Styles are returned.
; Return values .: ; Success: Integer or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bUserOnly not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bAppliedOnly not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Numbering Styles Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 0 = Success. No Numbering Styles found according to parameters.
;				   @Error 0 @Extended ? Return Array = Success. An Array containing all Numbering Styles matching the input parameters.
;				   +		@Extended contains the count of results returned. If Only a Document object is input, all available Numbering styles will be returned.
;				   +		Else if $bUserOnly is set to True, only User-Created Numbering Styles are returned.
;				   +		Else if $bAppliedOnly is set to True, only Applied Numbering Styles are returned.
;				   +		If Both are true then only User-Created Numbering styles that are applied are returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_NumStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_NumStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0
	Local $sExecute = ""
	Local $aStyles[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	Local $oStyles = $oDoc.StyleFamilies.getByName("NumberingStyles")
	If Not IsObj($oStyles) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	ReDim $aStyles[$oStyles.getCount()]

	If Not $bUserOnly And Not $bAppliedOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			$aStyles[$i] = $oStyles.getByIndex($i).Name() ; -- Can't use Display name due to special characters.
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
		Return SetError($__LO_STATUS_SUCCESS, $i, $aStyles)
	EndIf

	$sExecute = ($bUserOnly) ? ("($oStyles.getByIndex($i).isUserDefined())") : ($sExecute)
	$sExecute = ($bUserOnly And $bAppliedOnly) ? ($sExecute & " And ") : ($sExecute)
	$sExecute = ($bAppliedOnly) ? ($sExecute & "($oStyles.getByIndex($i).isInUse())") : ($sExecute)

	For $i = 0 To $oStyles.getCount() - 1
		If Execute($sExecute) Then
			$aStyles[$iCount] = $oStyles.getByIndex($i).Name() ; DisplayName
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	ReDim $aStyles[$iCount]

	Return ($iCount = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_SUCCESS, $iCount, $aStyles))
EndFunc   ;==>_LOWriter_NumStylesGetNames
